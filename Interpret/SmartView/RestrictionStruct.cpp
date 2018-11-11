#include <StdAfx.h>
#include <Interpret/SmartView/RestrictionStruct.h>
#include <Interpret/SmartView/PropertiesStruct.h>
#include <Interpret/String.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/SmartView/SmartView.h>

namespace smartview
{
	void RestrictionStruct::init(bool bRuleCondition, bool bExtended)
	{
		m_bRuleCondition = bRuleCondition;
		m_bExtended = bExtended;
	}

	void RestrictionStruct::Parse() { m_lpRes = BinToRestriction(0, m_bRuleCondition, m_bExtended); }

	void RestrictionStruct::ParseBlocks()
	{
		setRoot(L"Restriction:\r\n");
		ParseRestriction(m_lpRes, 0);
	}

	// Helper function for both RestrictionStruct and RuleConditionStruct
	// If bRuleCondition is true, parse restrictions as defined in [MS-OXCDATA] 2.12
	// If bRuleCondition is true, bExtendedCount controls whether the count fields in AND/OR restrictions is 16 or 32 bits
	//   https://msdn.microsoft.com/en-us/library/ee201126(v=exchg.80).aspx
	// If bRuleCondition is false, parse restrictions as defined in [MS-OXOCFG] 2.2.6.1.2
	// If bRuleCondition is false, ignore bExtendedCount (assumes true)
	//   https://msdn.microsoft.com/en-us/library/ee217813(v=exchg.80).aspx
	// Never fails, but will not parse restrictions above _MaxDepth
	// [MS-OXCDATA] 2.11.4 TaggedPropertyValue Structure
	SRestrictionStruct RestrictionStruct::BinToRestriction(ULONG ulDepth, bool bRuleCondition, bool bExtendedCount)
	{
		auto srRestriction = SRestrictionStruct{};

		if (bRuleCondition)
		{
			srRestriction.rt = m_Parser.Get<BYTE>();
		}
		else
		{
			srRestriction.rt = m_Parser.Get<DWORD>();
		}

		switch (srRestriction.rt)
		{
		case RES_AND:
			if (!bRuleCondition || bExtendedCount)
			{
				srRestriction.resAnd.cRes = m_Parser.Get<DWORD>();
			}
			else
			{
				srRestriction.resAnd.cRes = m_Parser.Get<WORD>();
			}

			if (srRestriction.resAnd.cRes && srRestriction.resAnd.cRes < _MaxEntriesExtraLarge && ulDepth < _MaxDepth)
			{
				srRestriction.resAnd.lpRes.reserve(srRestriction.resAnd.cRes);
				for (ULONG i = 0; i < srRestriction.resAnd.cRes; i++)
				{
					if (!m_Parser.RemainingBytes()) break;
					srRestriction.resAnd.lpRes.push_back(BinToRestriction(ulDepth + 1, bRuleCondition, bExtendedCount));
				}

				// TODO: Should we do this?
				//srRestriction.resAnd.lpRes.shrink_to_fit();
			}
			break;
		case RES_OR:
			if (!bRuleCondition || bExtendedCount)
			{
				srRestriction.resOr.cRes = m_Parser.Get<DWORD>();
			}
			else
			{
				srRestriction.resOr.cRes = m_Parser.Get<WORD>();
			}

			if (srRestriction.resOr.cRes && srRestriction.resOr.cRes < _MaxEntriesExtraLarge && ulDepth < _MaxDepth)
			{
				srRestriction.resOr.lpRes.reserve(srRestriction.resOr.cRes);
				for (ULONG i = 0; i < srRestriction.resOr.cRes; i++)
				{
					if (!m_Parser.RemainingBytes()) break;
					srRestriction.resOr.lpRes.push_back(BinToRestriction(ulDepth + 1, bRuleCondition, bExtendedCount));
				}
			}
			break;
		case RES_NOT:
			if (ulDepth < _MaxDepth && m_Parser.RemainingBytes())
			{
				srRestriction.resNot.lpRes.push_back(BinToRestriction(ulDepth + 1, bRuleCondition, bExtendedCount));
			}
			break;
		case RES_CONTENT:
			srRestriction.resContent.ulFuzzyLevel = m_Parser.Get<DWORD>();
			srRestriction.resContent.ulPropTag = m_Parser.Get<DWORD>();
			srRestriction.resContent.lpProp = BinToProps(1, bRuleCondition);
			break;
		case RES_PROPERTY:
			if (bRuleCondition)
				srRestriction.resProperty.relop = m_Parser.Get<BYTE>();
			else
				srRestriction.resProperty.relop = m_Parser.Get<DWORD>();

			srRestriction.resProperty.ulPropTag = m_Parser.Get<DWORD>();
			srRestriction.resProperty.lpProp = BinToProps(1, bRuleCondition);
			break;
		case RES_COMPAREPROPS:
			if (bRuleCondition)
				srRestriction.resCompareProps.relop = m_Parser.Get<BYTE>();
			else
				srRestriction.resCompareProps.relop = m_Parser.Get<DWORD>();

			srRestriction.resCompareProps.ulPropTag1 = m_Parser.Get<DWORD>();
			srRestriction.resCompareProps.ulPropTag2 = m_Parser.Get<DWORD>();
			break;
		case RES_BITMASK:
			if (bRuleCondition)
				srRestriction.resBitMask.relBMR = m_Parser.Get<BYTE>();
			else
				srRestriction.resBitMask.relBMR = m_Parser.Get<DWORD>();

			srRestriction.resBitMask.ulPropTag = m_Parser.Get<DWORD>();
			srRestriction.resBitMask.ulMask = m_Parser.Get<DWORD>();
			break;
		case RES_SIZE:
			if (bRuleCondition)
				srRestriction.resSize.relop = m_Parser.Get<BYTE>();
			else
				srRestriction.resSize.relop = m_Parser.Get<DWORD>();

			srRestriction.resSize.ulPropTag = m_Parser.Get<DWORD>();
			srRestriction.resSize.cb = m_Parser.Get<DWORD>();
			break;
		case RES_EXIST:
			srRestriction.resExist.ulPropTag = m_Parser.Get<DWORD>();
			break;
		case RES_SUBRESTRICTION:
			srRestriction.resSub.ulSubObject = m_Parser.Get<DWORD>();
			if (ulDepth < _MaxDepth && m_Parser.RemainingBytes())
			{
				srRestriction.resSub.lpRes.push_back(BinToRestriction(ulDepth + 1, bRuleCondition, bExtendedCount));
			}
			break;
		case RES_COMMENT:
			if (bRuleCondition)
				srRestriction.resComment.cValues = m_Parser.Get<BYTE>();
			else
				srRestriction.resComment.cValues = m_Parser.Get<DWORD>();

			srRestriction.resComment.lpProp = BinToProps(srRestriction.resComment.cValues, bRuleCondition);

			// Check if a restriction is present
			if (m_Parser.Get<BYTE>() && ulDepth < _MaxDepth && m_Parser.RemainingBytes())
			{
				srRestriction.resComment.lpRes.push_back(BinToRestriction(ulDepth + 1, bRuleCondition, bExtendedCount));
			}
			break;
		case RES_COUNT:
			srRestriction.resCount.ulCount = m_Parser.Get<DWORD>();
			if (ulDepth < _MaxDepth && m_Parser.RemainingBytes())
			{
				srRestriction.resCount.lpRes.push_back(BinToRestriction(ulDepth + 1, bRuleCondition, bExtendedCount));
			}
			break;
		}

		return srRestriction;
	}

	PropertiesStruct RestrictionStruct::BinToProps(DWORD cValues, bool bRuleCondition)
	{
		auto props = PropertiesStruct{};
		props.init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
		props.DisableJunkParsing();
		props.SetMaxEntries(cValues);
		if (bRuleCondition) props.EnableRuleConditionParsing();
		props.EnsureParsed();
		m_Parser.advance(props.GetCurrentOffset());

		return props;
	}

	// There may be restrictions with over 100 nested levels, but we're not going to try to parse them
#define _MaxRestrictionNesting 100

	void RestrictionStruct::ParseRestriction(_In_ const SRestrictionStruct& lpRes, ULONG ulTabLevel)
	{
		if (ulTabLevel > _MaxRestrictionNesting)
		{
			addHeader(L"Restriction nested too many (%d) levels.", _MaxRestrictionNesting);
			return;
		}

		std::wstring szTabs;
		for (ULONG i = 0; i < ulTabLevel; i++)
		{
			szTabs += L"\t"; // STRING_OK
		}

		std::wstring szPropNum;
		auto szFlags = interpretprop::InterpretFlags(flagRestrictionType, lpRes.rt);
		addBlock(
			lpRes.rt, L"%1!ws!lpRes->rt = 0x%2!X! = %3!ws!\r\n", szTabs.c_str(), lpRes.rt.getData(), szFlags.c_str());

		auto i = 0;
		switch (lpRes.rt)
		{
		case RES_COMPAREPROPS:
			szFlags = interpretprop::InterpretFlags(flagRelop, lpRes.resCompareProps.relop);
			addBlock(
				lpRes.resCompareProps.relop,
				L"%1!ws!lpRes->res.resCompareProps.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes.resCompareProps.relop.getData());
			addBlock(
				lpRes.resCompareProps.ulPropTag1,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag1 = %2!ws!\r\n",
				szTabs.c_str(),
				interpretprop::TagToString(lpRes.resCompareProps.ulPropTag1, nullptr, false, true).c_str());
			addBlock(
				lpRes.resCompareProps.ulPropTag2,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag2 = %2!ws!\r\n",
				szTabs.c_str(),
				interpretprop::TagToString(lpRes.resCompareProps.ulPropTag2, nullptr, false, true).c_str());
			break;
		case RES_AND:
			addBlock(
				lpRes.resAnd.cRes,
				L"%1!ws!lpRes->res.resAnd.cRes = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resAnd.cRes.getData());
			for (auto res : lpRes.resAnd.lpRes)
			{
				addHeader(L"%1!ws!lpRes->res.resAnd.lpRes[0x%2!08X!]\r\n", szTabs.c_str(), i++);
				ParseRestriction(res, ulTabLevel + 1);
			}
			break;
		case RES_OR:
			addBlock(
				lpRes.resOr.cRes,
				L"%1!ws!lpRes->res.resOr.cRes = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resOr.cRes.getData());
			for (auto res : lpRes.resOr.lpRes)
			{
				addHeader(L"%1!ws!lpRes->res.resOr.lpRes[0x%2!08X!]\r\n", szTabs.c_str(), i++);
				ParseRestriction(res, ulTabLevel + 1);
			}
			break;
		case RES_NOT:
			addBlock(
				lpRes.resNot.ulReserved,
				L"%1!ws!lpRes->res.resNot.ulReserved = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resNot.ulReserved.getData());
			addHeader(L"%1!ws!lpRes->res.resNot.lpRes\r\n", szTabs.c_str());

			// There should only be one sub restriction here
			if (lpRes.resNot.lpRes.size())
			{
				ParseRestriction(lpRes.resNot.lpRes[0], ulTabLevel + 1);
			}
			break;
		case RES_COUNT:
			addBlock(
				lpRes.resCount.ulCount,
				L"%1!ws!lpRes->res.resCount.ulCount = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resCount.ulCount.getData());
			addHeader(L"%1!ws!lpRes->res.resCount.lpRes\r\n", szTabs.c_str());

			if (lpRes.resCount.lpRes.size())
			{
				ParseRestriction(lpRes.resCount.lpRes[0], ulTabLevel + 1);
			}
			break;
		case RES_CONTENT:
			szFlags = interpretprop::InterpretFlags(flagFuzzyLevel, lpRes.resContent.ulFuzzyLevel);
			addBlock(
				lpRes.resContent.ulFuzzyLevel,
				L"%1!ws!lpRes->res.resContent.ulFuzzyLevel = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes.resContent.ulFuzzyLevel.getData());
			addBlock(
				lpRes.resContent.ulPropTag,
				L"%1!ws!lpRes->res.resContent.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				interpretprop::TagToString(lpRes.resContent.ulPropTag, nullptr, false, true).c_str());

			if (lpRes.resContent.lpProp.Props().size())
			{
				addBlock(
					lpRes.resContent.lpProp.Props()[0].ulPropTag,
					L"%1!ws!lpRes->res.resContent.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					interpretprop::TagToString(lpRes.resContent.lpProp.Props()[0].ulPropTag, nullptr, false, true)
						.c_str());
				// TODO: use blocks
				addHeader(
					L"%1!ws!lpRes->res.resContent.lpProp->Value = %2!ws!\r\n",
					szTabs.c_str(),
					lpRes.resContent.lpProp.Props()[0].PropString().c_str());
				addHeader(
					L"%1!ws!\tAlt: %2!ws!", szTabs.c_str(), lpRes.resContent.lpProp.Props()[0].AltPropString().c_str());
				addBlankLine();
			}
			break;
		case RES_PROPERTY:
			szFlags = interpretprop::InterpretFlags(flagRelop, lpRes.resProperty.relop);
			addBlock(
				lpRes.resProperty.relop,
				L"%1!ws!lpRes->res.resProperty.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes.resProperty.relop.getData());
			addBlock(
				lpRes.resProperty.ulPropTag,
				L"%1!ws!lpRes->res.resProperty.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				interpretprop::TagToString(lpRes.resProperty.ulPropTag, nullptr, false, true).c_str());
			if (lpRes.resProperty.lpProp.Props().size())
			{
				addBlock(
					lpRes.resProperty.lpProp.Props()[0].ulPropTag,
					L"%1!ws!lpRes->res.resProperty.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					interpretprop::TagToString(lpRes.resProperty.lpProp.Props()[0].ulPropTag, nullptr, false, true)
						.c_str());
				// TODO: use blocks
				addHeader(
					L"%1!ws!lpRes->res.resProperty.lpProp->Value = %2!ws!\r\n",
					szTabs.c_str(),
					lpRes.resProperty.lpProp.Props()[0].PropString().c_str());
				addHeader(
					L"%1!ws!\tAlt: %2!ws!\r\n",
					szTabs.c_str(),
					lpRes.resProperty.lpProp.Props()[0].AltPropString().c_str());
				szPropNum = lpRes.resProperty.lpProp.Props()[0].PropNum();
				if (!szPropNum.empty())
				{
					// TODO: use block
					addHeader(L"%1!ws!\tFlags: %2!ws!", szTabs.c_str(), szPropNum.c_str());
				}
			}
			break;
		case RES_BITMASK:
			szFlags = interpretprop::InterpretFlags(flagBitmask, lpRes.resBitMask.relBMR);
			addBlock(
				lpRes.resBitMask.relBMR,
				L"%1!ws!lpRes->res.resBitMask.relBMR = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes.resBitMask.relBMR.getData());
			addBlock(
				lpRes.resBitMask.ulMask,
				L"%1!ws!lpRes->res.resBitMask.ulMask = 0x%2!08X!",
				szTabs.c_str(),
				lpRes.resBitMask.ulMask.getData());
			szPropNum = InterpretNumberAsStringProp(lpRes.resBitMask.ulMask, lpRes.resBitMask.ulPropTag);
			if (!szPropNum.empty())
			{
				addBlock(lpRes.resBitMask.ulMask, L": %1!ws!", szPropNum.c_str());
			}

			addBlankLine();
			addBlock(
				lpRes.resBitMask.ulPropTag,
				L"%1!ws!lpRes->res.resBitMask.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				interpretprop::TagToString(lpRes.resBitMask.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_SIZE:
			szFlags = interpretprop::InterpretFlags(flagRelop, lpRes.resSize.relop);
			addBlock(
				lpRes.resSize.relop,
				L"%1!ws!lpRes->res.resSize.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes.resSize.relop.getData());
			addBlock(
				lpRes.resSize.cb,
				L"%1!ws!lpRes->res.resSize.cb = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resSize.cb.getData());
			addBlock(
				lpRes.resSize.ulPropTag,
				L"%1!ws!lpRes->res.resSize.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				interpretprop::TagToString(lpRes.resSize.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_EXIST:
			addBlock(
				lpRes.resExist.ulPropTag,
				L"%1!ws!lpRes->res.resExist.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				interpretprop::TagToString(lpRes.resExist.ulPropTag, nullptr, false, true).c_str());
			addBlock(
				lpRes.resExist.ulReserved1,
				L"%1!ws!lpRes->res.resExist.ulReserved1 = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resExist.ulReserved1.getData());
			addBlock(
				lpRes.resExist.ulReserved2,
				L"%1!ws!lpRes->res.resExist.ulReserved2 = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resExist.ulReserved2.getData());
			break;
		case RES_SUBRESTRICTION:
			addBlock(
				lpRes.resSub.ulSubObject,
				L"%1!ws!lpRes->res.resSub.ulSubObject = %2!ws!\r\n",
				szTabs.c_str(),
				interpretprop::TagToString(lpRes.resSub.ulSubObject, nullptr, false, true).c_str());
			addHeader(L"%1!ws!lpRes->res.resSub.lpRes\r\n", szTabs.c_str());

			if (lpRes.resSub.lpRes.size())
			{
				ParseRestriction(lpRes.resSub.lpRes[0], ulTabLevel + 1);
			}
			break;
		case RES_COMMENT:
			addBlock(
				lpRes.resComment.cValues,
				L"%1!ws!lpRes->res.resComment.cValues = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resComment.cValues.getData());
			for (auto prop : lpRes.resComment.lpProp.Props())
			{
				addBlock(
					prop.ulPropTag,
					L"%1!ws!lpRes->res.resComment.lpProp[0x%2!08X!].ulPropTag = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					interpretprop::TagToString(prop.ulPropTag, nullptr, false, true).c_str());
				// TODO: use blocks
				addHeader(
					L"%1!ws!lpRes->res.resComment.lpProp[0x%2!08X!].Value = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					prop.PropString().c_str());
				addHeader(L"%1!ws!\tAlt: %2!ws!\r\n", szTabs.c_str(), prop.AltPropString().c_str());
				i++;
			}

			addHeader(L"%1!ws!lpRes->res.resComment.lpRes\r\n", szTabs.c_str());
			if (lpRes.resComment.lpRes.size())
			{
				ParseRestriction(lpRes.resComment.lpRes[0], ulTabLevel + 1);
			}
			break;
		case RES_ANNOTATION:
			addBlock(
				lpRes.resComment.cValues,
				L"%1!ws!lpRes->res.resAnnotation.cValues = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resComment.cValues.getData());
			for (auto prop : lpRes.resComment.lpProp.Props())
			{
				addBlock(
					prop.ulPropTag,
					L"%1!ws!lpRes->res.resAnnotation.lpProp[0x%2!08X!].ulPropTag = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					interpretprop::TagToString(prop.ulPropTag, nullptr, false, true).c_str());
				// TODO: Use blocks
				addHeader(
					L"%1!ws!lpRes->res.resAnnotation.lpProp[0x%2!08X!].Value = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					prop.PropString().c_str());
				addHeader(L"%1!ws!\tAlt: %2!ws!\r\n", szTabs.c_str(), prop.AltPropString().c_str());
				i++;
			}

			addHeader(L"%1!ws!lpRes->res.resAnnotation.lpRes\r\n", szTabs.c_str());

			if (lpRes.resComment.lpRes.size())
			{
				ParseRestriction(lpRes.resComment.lpRes[0], ulTabLevel + 1);
			}
			break;
		}
	}
} // namespace smartview