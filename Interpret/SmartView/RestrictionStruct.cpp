#include "stdafx.h"
#include "RestrictionStruct.h"
#include "PropertyStruct.h"
#include <Interpret/String.h>
#include <Interpret/InterpretProp.h>

RestrictionStruct::RestrictionStruct()
{
	m_bRuleCondition = false;
	m_bExtended = false;
	m_lpRes = nullptr;
}

void RestrictionStruct::Init(bool bRuleCondition, bool bExtended)
{
	m_bRuleCondition = bRuleCondition;
	m_bExtended = bExtended;
}

void RestrictionStruct::Parse()
{
	m_lpRes = reinterpret_cast<LPSRestriction>(Allocate(sizeof SRestriction));

	if (m_lpRes)
	{
		(void)BinToRestriction(
			0,
			m_lpRes,
			m_bRuleCondition,
			m_bExtended);
	}
}

_Check_return_ std::wstring RestrictionStruct::ToStringInternal()
{
	std::wstring szRestriction;

	szRestriction = strings::formatmessage(IDS_RESTRICTIONDATA);
	szRestriction += RestrictionToString(m_lpRes, nullptr);

	return szRestriction;
}

// Helper function for both RestrictionStruct and RuleConditionStruct
// If bRuleCondition is true, parse restrictions defined in [MS-OXCDATA] 2.13
// If bRuleCondition is true, bExtendedCount controls whether the count fields in AND/OR restrictions is 16 or 32 bits
// If bRuleCondition is false, parse restrictions defined in [MS-OXOCFG] 2.2.4.1.2
// If bRuleCondition is false, ignore bExtendedCount (assumes true)
// Returns true if it successfully read a restriction
bool RestrictionStruct::BinToRestriction(ULONG ulDepth, _In_ LPSRestriction psrRestriction, bool bRuleCondition, bool bExtendedCount)
{
	if (ulDepth > _MaxDepth) return false;
	if (!psrRestriction) return false;

	auto bRet = true;
	DWORD dwTemp = 0;

	if (bRuleCondition)
	{
		psrRestriction->rt = m_Parser.Get<BYTE>();;
	}
	else
	{
		psrRestriction->rt = m_Parser.Get<DWORD>();
	}

	switch (psrRestriction->rt)
	{
	case RES_AND:
	case RES_OR:
		if (!bRuleCondition || bExtendedCount)
		{
			dwTemp = m_Parser.Get<DWORD>();
		}
		else
		{
			dwTemp = m_Parser.Get<WORD>();
		}

		psrRestriction->res.resAnd.cRes = dwTemp;
		if (psrRestriction->res.resAnd.cRes &&
			psrRestriction->res.resAnd.cRes < _MaxEntriesExtraLarge)
		{
			psrRestriction->res.resAnd.lpRes = reinterpret_cast<SRestriction*>(AllocateArray(psrRestriction->res.resAnd.cRes, sizeof SRestriction));
			if (psrRestriction->res.resAnd.lpRes)
			{
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
		psrRestriction->res.resNot.lpRes = reinterpret_cast<LPSRestriction>(Allocate(sizeof SRestriction));
		if (psrRestriction->res.resNot.lpRes)
		{
			bRet = BinToRestriction(
				ulDepth + 1,
				psrRestriction->res.resNot.lpRes,
				bRuleCondition,
				bExtendedCount);
		}
		break;
	case RES_CONTENT:
		psrRestriction->res.resContent.ulFuzzyLevel = m_Parser.Get<DWORD>();
		psrRestriction->res.resContent.ulPropTag = m_Parser.Get<DWORD>();
		psrRestriction->res.resContent.lpProp = BinToSPropValue(
			1,
			bRuleCondition);
		break;
	case RES_PROPERTY:
		if (bRuleCondition)
			psrRestriction->res.resProperty.relop = m_Parser.Get<BYTE>();
		else
			psrRestriction->res.resProperty.relop = m_Parser.Get<DWORD>();

		psrRestriction->res.resProperty.ulPropTag = m_Parser.Get<DWORD>();
		psrRestriction->res.resProperty.lpProp = BinToSPropValue(
			1,
			bRuleCondition);
		break;
	case RES_COMPAREPROPS:
		if (bRuleCondition)
			psrRestriction->res.resCompareProps.relop = m_Parser.Get<BYTE>();
		else
			psrRestriction->res.resCompareProps.relop = m_Parser.Get<DWORD>();

		psrRestriction->res.resCompareProps.ulPropTag1 = m_Parser.Get<DWORD>();
		psrRestriction->res.resCompareProps.ulPropTag2 = m_Parser.Get<DWORD>();
		break;
	case RES_BITMASK:
		if (bRuleCondition)
			psrRestriction->res.resBitMask.relBMR = m_Parser.Get<BYTE>();
		else
			psrRestriction->res.resBitMask.relBMR = m_Parser.Get<DWORD>();

		psrRestriction->res.resBitMask.ulPropTag = m_Parser.Get<DWORD>();
		psrRestriction->res.resBitMask.ulMask = m_Parser.Get<DWORD>();
		break;
	case RES_SIZE:
		if (bRuleCondition)
			psrRestriction->res.resSize.relop = m_Parser.Get<BYTE>();
		else
			psrRestriction->res.resSize.relop = m_Parser.Get<DWORD>();

		psrRestriction->res.resSize.ulPropTag = m_Parser.Get<DWORD>();
		psrRestriction->res.resSize.cb = m_Parser.Get<DWORD>();
		break;
	case RES_EXIST:
		psrRestriction->res.resExist.ulPropTag = m_Parser.Get<DWORD>();
		break;
	case RES_SUBRESTRICTION:
		psrRestriction->res.resSub.ulSubObject = m_Parser.Get<DWORD>();
		psrRestriction->res.resSub.lpRes = reinterpret_cast<LPSRestriction>(Allocate(sizeof SRestriction));
		if (psrRestriction->res.resSub.lpRes)
		{
			bRet = BinToRestriction(
				ulDepth + 1,
				psrRestriction->res.resSub.lpRes,
				bRuleCondition,
				bExtendedCount);
		}
		break;
	case RES_COMMENT:
		if (bRuleCondition)
			psrRestriction->res.resComment.cValues = m_Parser.Get<BYTE>();
		else
			psrRestriction->res.resComment.cValues = m_Parser.Get<DWORD>();

		psrRestriction->res.resComment.lpProp = BinToSPropValue(
			psrRestriction->res.resComment.cValues,
			bRuleCondition);

		// Check if a restriction is present
		if (m_Parser.Get<BYTE>())
		{
			psrRestriction->res.resComment.lpRes = reinterpret_cast<LPSRestriction>(Allocate(sizeof SRestriction));
			if (psrRestriction->res.resComment.lpRes)
			{
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
		psrRestriction->res.resNot.ulReserved = m_Parser.Get<DWORD>();
		psrRestriction->res.resNot.lpRes = reinterpret_cast<LPSRestriction>(Allocate(sizeof SRestriction));
		if (psrRestriction->res.resNot.lpRes)
		{
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