#include "stdafx.h"
#include "..\stdafx.h"
#include "RestrictionStruct.h"
#include "SmartView.h"
#include "..\String.h"
#include "..\InterpretProp.h"

RestrictionStruct::RestrictionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_lpRes = NULL;
}

RestrictionStruct::~RestrictionStruct()
{
	DeleteRestriction(m_lpRes);
	delete[] m_lpRes;
}

void RestrictionStruct::Parse()
{
	m_lpRes = new SRestriction;

	if (m_lpRes)
	{
		size_t cbBytesRead = 0;
		memset(m_lpRes, 0, sizeof(SRestriction));

		(void)BinToRestriction(
			0,
			m_Parser.RemainingBytes(),
			m_Parser.GetCurrentAddress(),
			&cbBytesRead,
			m_lpRes,
			false,
			true);
		m_Parser.Advance(cbBytesRead);
	}
}

_Check_return_ wstring RestrictionStruct::ToStringInternal()
{
	wstring szRestriction;

	szRestriction = formatmessage(IDS_RESTRICTIONDATA);
	szRestriction += LPTSTRToWstring((LPTSTR)(LPCTSTR)RestrictionToString(m_lpRes, NULL));

	return szRestriction;
}

// Helper function for both RestrictionStruct and RuleConditionStruct
// Caller allocates with new. Clean up with DeleteRestriction and delete[].
// If bRuleCondition is true, parse restrictions defined in [MS-OXCDATA] 2.13
// If bRuleCondition is true, bExtendedCount controls whether the count fields in AND/OR restrictions is 16 or 32 bits
// If bRuleCondition is false, parse restrictions defined in [MS-OXOCFG] 2.2.4.1.2
// If bRuleCondition is false, ignore bExtendedCount (assumes true)
// Returns true if it successfully read a restriction
bool BinToRestriction(ULONG ulDepth, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _In_ LPSRestriction psrRestriction, bool bRuleCondition, bool bExtendedCount)
{
	if (ulDepth > _MaxDepth) return false;
	if (!lpBin) return false;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	if (!psrRestriction) return false;
	CBinaryParser Parser(cbBin, lpBin);

	BYTE bTemp = 0;
	WORD wTemp = 0;
	DWORD dwTemp = 0;
	DWORD i = 0;
	size_t cbOffset = 0;
	size_t cbBytesRead = 0;
	bool bRet = true;

	if (bRuleCondition)
	{
		Parser.GetBYTE(&bTemp);
		psrRestriction->rt = bTemp;
	}
	else
	{
		Parser.GetDWORD(&dwTemp);
		psrRestriction->rt = dwTemp;
	}

	switch (psrRestriction->rt)
	{
	case RES_AND:
	case RES_OR:
		if (!bRuleCondition || bExtendedCount)
		{
			Parser.GetDWORD(&dwTemp);
		}
		else
		{
			Parser.GetWORD(&wTemp);
			dwTemp = wTemp;
		}
		psrRestriction->res.resAnd.cRes = dwTemp;
		if (psrRestriction->res.resAnd.cRes &&
			psrRestriction->res.resAnd.cRes < _MaxEntriesExtraLarge)
		{
			psrRestriction->res.resAnd.lpRes = new SRestriction[psrRestriction->res.resAnd.cRes];
			if (psrRestriction->res.resAnd.lpRes)
			{
				memset(psrRestriction->res.resAnd.lpRes, 0, sizeof(SRestriction)* psrRestriction->res.resAnd.cRes);
				for (i = 0; i < psrRestriction->res.resAnd.cRes; i++)
				{
					cbOffset = Parser.GetCurrentOffset();
					bRet = BinToRestriction(
						ulDepth + 1,
						(ULONG)Parser.RemainingBytes(),
						lpBin + cbOffset,
						&cbBytesRead,
						&psrRestriction->res.resAnd.lpRes[i],
						bRuleCondition,
						bExtendedCount);
					Parser.Advance(cbBytesRead);
					if (!bRet) break;
				}
			}
		}
		break;
	case RES_NOT:
		psrRestriction->res.resNot.lpRes = new SRestriction;
		if (psrRestriction->res.resNot.lpRes)
		{
			memset(psrRestriction->res.resNot.lpRes, 0, sizeof(SRestriction));
			cbOffset = Parser.GetCurrentOffset();
			bRet = BinToRestriction(
				ulDepth + 1,
				(ULONG)Parser.RemainingBytes(),
				lpBin + cbOffset,
				&cbBytesRead,
				psrRestriction->res.resNot.lpRes,
				bRuleCondition,
				bExtendedCount);
			Parser.Advance(cbBytesRead);
		}
		break;
	case RES_CONTENT:
		Parser.GetDWORD(&psrRestriction->res.resContent.ulFuzzyLevel);
		Parser.GetDWORD(&psrRestriction->res.resContent.ulPropTag);
		cbOffset = Parser.GetCurrentOffset();
		psrRestriction->res.resContent.lpProp = BinToSPropValue(
			(ULONG)Parser.RemainingBytes(),
			lpBin + cbOffset,
			1,
			&cbBytesRead,
			bRuleCondition);
		Parser.Advance(cbBytesRead);
		break;
	case RES_PROPERTY:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resProperty.relop);
		else
			Parser.GetDWORD(&psrRestriction->res.resProperty.relop);
		Parser.GetDWORD(&psrRestriction->res.resProperty.ulPropTag);
		cbOffset = Parser.GetCurrentOffset();
		psrRestriction->res.resProperty.lpProp = BinToSPropValue(
			(ULONG)Parser.RemainingBytes(),
			lpBin + cbOffset,
			1,
			&cbBytesRead,
			bRuleCondition);
		Parser.Advance(cbBytesRead);
		break;
	case RES_COMPAREPROPS:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resCompareProps.relop);
		else
			Parser.GetDWORD(&psrRestriction->res.resCompareProps.relop);
		Parser.GetDWORD(&psrRestriction->res.resCompareProps.ulPropTag1);
		Parser.GetDWORD(&psrRestriction->res.resCompareProps.ulPropTag2);
		break;
	case RES_BITMASK:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resBitMask.relBMR);
		else
			Parser.GetDWORD(&psrRestriction->res.resBitMask.relBMR);
		Parser.GetDWORD(&psrRestriction->res.resBitMask.ulPropTag);
		Parser.GetDWORD(&psrRestriction->res.resBitMask.ulMask);
		break;
	case RES_SIZE:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resSize.relop);
		else
			Parser.GetDWORD(&psrRestriction->res.resSize.relop);
		Parser.GetDWORD(&psrRestriction->res.resSize.ulPropTag);
		Parser.GetDWORD(&psrRestriction->res.resSize.cb);
		break;
	case RES_EXIST:
		Parser.GetDWORD(&psrRestriction->res.resExist.ulPropTag);
		break;
	case RES_SUBRESTRICTION:
		Parser.GetDWORD(&psrRestriction->res.resSub.ulSubObject);
		psrRestriction->res.resSub.lpRes = new SRestriction;
		if (psrRestriction->res.resSub.lpRes)
		{
			memset(psrRestriction->res.resSub.lpRes, 0, sizeof(SRestriction));
			cbOffset = Parser.GetCurrentOffset();
			bRet = BinToRestriction(
				ulDepth + 1,
				(ULONG)Parser.RemainingBytes(),
				lpBin + cbOffset,
				&cbBytesRead,
				psrRestriction->res.resSub.lpRes,
				bRuleCondition,
				bExtendedCount);
			Parser.Advance(cbBytesRead);
		}
		break;
	case RES_COMMENT:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resComment.cValues);
		else
			Parser.GetDWORD(&psrRestriction->res.resComment.cValues);
		cbOffset = Parser.GetCurrentOffset();
		psrRestriction->res.resProperty.lpProp = BinToSPropValue(
			(ULONG)Parser.RemainingBytes(),
			lpBin + cbOffset,
			psrRestriction->res.resComment.cValues,
			&cbBytesRead,
			bRuleCondition);
		Parser.Advance(cbBytesRead);

		// Check if a restriction is present
		Parser.GetBYTE(&bTemp);
		if (bTemp)
		{
			psrRestriction->res.resComment.lpRes = new SRestriction;
			if (psrRestriction->res.resComment.lpRes)
			{
				memset(psrRestriction->res.resComment.lpRes, 0, sizeof(SRestriction));
				cbOffset = Parser.GetCurrentOffset();
				bRet = BinToRestriction(
					ulDepth + 1,
					(ULONG)Parser.RemainingBytes(),
					lpBin + cbOffset,
					&cbBytesRead,
					psrRestriction->res.resComment.lpRes,
					bRuleCondition,
					bExtendedCount);
				Parser.Advance(cbBytesRead);
			}
		}
		break;
	case RES_COUNT:
		// RES_COUNT and RES_NOT look the same, so we use the resNot member here
		Parser.GetDWORD(&psrRestriction->res.resNot.ulReserved);
		psrRestriction->res.resNot.lpRes = new SRestriction;
		if (psrRestriction->res.resNot.lpRes)
		{
			memset(psrRestriction->res.resNot.lpRes, 0, sizeof(SRestriction));
			cbOffset = Parser.GetCurrentOffset();
			bRet = BinToRestriction(
				ulDepth + 1,
				(ULONG)Parser.RemainingBytes(),
				lpBin + cbOffset,
				&cbBytesRead,
				psrRestriction->res.resNot.lpRes,
				bRuleCondition,
				bExtendedCount);
			Parser.Advance(cbBytesRead);
		}
		break;
	}

	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();
	return bRet;
}

// Does not delete lpRes, which must be released manually. See RES_AND case below.
void DeleteRestriction(_In_ LPSRestriction lpRes)
{
	if (!lpRes) return;

	switch (lpRes->rt)
	{
	case RES_AND:
	case RES_OR:
		if (lpRes->res.resAnd.lpRes)
		{
			DWORD i = 0;
			for (i = 0; i < lpRes->res.resAnd.cRes; i++)
			{
				DeleteRestriction(&lpRes->res.resAnd.lpRes[i]);
			}
		}

		delete[] lpRes->res.resAnd.lpRes;
		break;
		// Structures for these two types are identical
	case RES_NOT:
	case RES_COUNT:
		DeleteRestriction(lpRes->res.resNot.lpRes);
		delete[] lpRes->res.resNot.lpRes;
		break;
	case RES_CONTENT:
		DeleteSPropVal(1, lpRes->res.resContent.lpProp);
		break;
	case RES_PROPERTY:
		DeleteSPropVal(1, lpRes->res.resProperty.lpProp);
		break;
	case RES_SUBRESTRICTION:
		DeleteRestriction(lpRes->res.resSub.lpRes);
		delete[] lpRes->res.resSub.lpRes;
		break;
	case RES_COMMENT:
		if (lpRes->res.resComment.cValues)
		{
			DeleteSPropVal(lpRes->res.resComment.cValues, lpRes->res.resComment.lpProp);
		}

		DeleteRestriction(lpRes->res.resComment.lpRes);
		delete[] lpRes->res.resComment.lpRes;
		break;
	}
}