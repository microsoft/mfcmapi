#include "stdafx.h"
#include "RestrictionStruct.h"
#include "PropertyStruct.h"
#include "String.h"
#include "InterpretProp.h"

RestrictionStruct::RestrictionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, bool bRuleCondition, bool bExtended) : SmartViewParser(cbBin, lpBin)
{
	m_bRuleCondition = bRuleCondition;
	m_bExtended = bExtended;
	m_lpRes = nullptr;
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
		memset(m_lpRes, 0, sizeof(SRestriction));

		(void)BinToRestriction(
			0,
			m_lpRes,
			m_bRuleCondition,
			m_bExtended);
	}
}

_Check_return_ wstring RestrictionStruct::ToStringInternal()
{
	wstring szRestriction;

	szRestriction = formatmessage(IDS_RESTRICTIONDATA);
	szRestriction += RestrictionToString(m_lpRes, nullptr);

	return szRestriction;
}

// Helper function for both RestrictionStruct and RuleConditionStruct
// Caller allocates with new. Clean up with DeleteRestriction and delete[].
// If bRuleCondition is true, parse restrictions defined in [MS-OXCDATA] 2.13
// If bRuleCondition is true, bExtendedCount controls whether the count fields in AND/OR restrictions is 16 or 32 bits
// If bRuleCondition is false, parse restrictions defined in [MS-OXOCFG] 2.2.4.1.2
// If bRuleCondition is false, ignore bExtendedCount (assumes true)
// Returns true if it successfully read a restriction
bool RestrictionStruct::BinToRestriction(ULONG ulDepth, _In_ LPSRestriction psrRestriction, bool bRuleCondition, bool bExtendedCount)
{
	if (ulDepth > _MaxDepth) return false;
	if (!psrRestriction) return false;

	BYTE bTemp = 0;
	WORD wTemp = 0;
	DWORD dwTemp = 0;
	auto bRet = true;

	if (bRuleCondition)
	{
		m_Parser.GetBYTE(&bTemp);
		psrRestriction->rt = bTemp;
	}
	else
	{
		m_Parser.GetDWORD(&dwTemp);
		psrRestriction->rt = dwTemp;
	}

	switch (psrRestriction->rt)
	{
	case RES_AND:
	case RES_OR:
		if (!bRuleCondition || bExtendedCount)
		{
			m_Parser.GetDWORD(&dwTemp);
		}
		else
		{
			m_Parser.GetWORD(&wTemp);
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
				for (ULONG i = 0; i < psrRestriction->res.resAnd.cRes; i++)
				{
					bRet = BinToRestriction(
						ulDepth + 1,
						&psrRestriction->res.resAnd.lpRes[i],
						bRuleCondition,
						bExtendedCount);
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
			bRet = BinToRestriction(
				ulDepth + 1,
				psrRestriction->res.resNot.lpRes,
				bRuleCondition,
				bExtendedCount);
		}
		break;
	case RES_CONTENT:
		m_Parser.GetDWORD(&psrRestriction->res.resContent.ulFuzzyLevel);
		m_Parser.GetDWORD(&psrRestriction->res.resContent.ulPropTag);
		psrRestriction->res.resContent.lpProp = BinToSPropValue(
			1,
			bRuleCondition);
		break;
	case RES_PROPERTY:
		if (bRuleCondition)
			m_Parser.GetBYTE(reinterpret_cast<LPBYTE>(&psrRestriction->res.resProperty.relop));
		else
			m_Parser.GetDWORD(&psrRestriction->res.resProperty.relop);

		m_Parser.GetDWORD(&psrRestriction->res.resProperty.ulPropTag);
		psrRestriction->res.resProperty.lpProp = BinToSPropValue(
			1,
			bRuleCondition);
		break;
	case RES_COMPAREPROPS:
		if (bRuleCondition)
			m_Parser.GetBYTE(reinterpret_cast<LPBYTE>(&psrRestriction->res.resCompareProps.relop));
		else
			m_Parser.GetDWORD(&psrRestriction->res.resCompareProps.relop);

		m_Parser.GetDWORD(&psrRestriction->res.resCompareProps.ulPropTag1);
		m_Parser.GetDWORD(&psrRestriction->res.resCompareProps.ulPropTag2);
		break;
	case RES_BITMASK:
		if (bRuleCondition)
			m_Parser.GetBYTE(reinterpret_cast<LPBYTE>(&psrRestriction->res.resBitMask.relBMR));
		else
			m_Parser.GetDWORD(&psrRestriction->res.resBitMask.relBMR);

		m_Parser.GetDWORD(&psrRestriction->res.resBitMask.ulPropTag);
		m_Parser.GetDWORD(&psrRestriction->res.resBitMask.ulMask);
		break;
	case RES_SIZE:
		if (bRuleCondition)
			m_Parser.GetBYTE(reinterpret_cast<LPBYTE>(&psrRestriction->res.resSize.relop));
		else
			m_Parser.GetDWORD(&psrRestriction->res.resSize.relop);

		m_Parser.GetDWORD(&psrRestriction->res.resSize.ulPropTag);
		m_Parser.GetDWORD(&psrRestriction->res.resSize.cb);
		break;
	case RES_EXIST:
		m_Parser.GetDWORD(&psrRestriction->res.resExist.ulPropTag);
		break;
	case RES_SUBRESTRICTION:
		m_Parser.GetDWORD(&psrRestriction->res.resSub.ulSubObject);
		psrRestriction->res.resSub.lpRes = new SRestriction;
		if (psrRestriction->res.resSub.lpRes)
		{
			memset(psrRestriction->res.resSub.lpRes, 0, sizeof(SRestriction));
			bRet = BinToRestriction(
				ulDepth + 1,
				psrRestriction->res.resSub.lpRes,
				bRuleCondition,
				bExtendedCount);
		}
		break;
	case RES_COMMENT:
		if (bRuleCondition)
			m_Parser.GetBYTE(reinterpret_cast<LPBYTE>(&psrRestriction->res.resComment.cValues));
		else
			m_Parser.GetDWORD(&psrRestriction->res.resComment.cValues);

		psrRestriction->res.resProperty.lpProp = BinToSPropValue(
			psrRestriction->res.resComment.cValues,
			bRuleCondition);

		// Check if a restriction is present
		m_Parser.GetBYTE(&bTemp);
		if (bTemp)
		{
			psrRestriction->res.resComment.lpRes = new SRestriction;
			if (psrRestriction->res.resComment.lpRes)
			{
				memset(psrRestriction->res.resComment.lpRes, 0, sizeof(SRestriction));
				bRet = BinToRestriction(
					ulDepth + 1,
					psrRestriction->res.resComment.lpRes,
					bRuleCondition,
					bExtendedCount);
			}
		}
		break;
	case RES_COUNT:
		// RES_COUNT and RES_NOT look the same, so we use the resNot member here
		m_Parser.GetDWORD(&psrRestriction->res.resNot.ulReserved);
		psrRestriction->res.resNot.lpRes = new SRestriction;
		if (psrRestriction->res.resNot.lpRes)
		{
			memset(psrRestriction->res.resNot.lpRes, 0, sizeof(SRestriction));
			bRet = BinToRestriction(
				ulDepth + 1,
				psrRestriction->res.resNot.lpRes,
				bRuleCondition,
				bExtendedCount);
		}
		break;
	}

	return bRet;
}

// Does not delete lpRes, which must be released manually. See RES_AND case below.
void RestrictionStruct::DeleteRestriction(_In_ LPSRestriction lpRes) const
{
	if (!lpRes) return;

	switch (lpRes->rt)
	{
	case RES_AND:
	case RES_OR:
		if (lpRes->res.resAnd.lpRes)
		{
			for (ULONG i = 0; i < lpRes->res.resAnd.cRes; i++)
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
