#include <core/stdafx.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	void RestrictionStruct::init(bool bRuleCondition, bool bExtended)
	{
		m_bRuleCondition = bRuleCondition;
		m_bExtended = bExtended;
	}

	void RestrictionStruct::Parse() { m_lpRes.init(m_Parser, 0, m_bRuleCondition, m_bExtended); }

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
	void SRestrictionStruct::init(
		std::shared_ptr<binaryParser> parser,
		ULONG ulDepth,
		bool bRuleCondition,
		bool bExtendedCount)
	{
		if (bRuleCondition)
		{
			rt = parser->Get<BYTE>();
		}
		else
		{
			rt = parser->Get<DWORD>();
		}

		switch (rt)
		{
		case RES_AND:
			if (!bRuleCondition || bExtendedCount)
			{
				resAnd.cRes = parser->Get<DWORD>();
			}
			else
			{
				resAnd.cRes = parser->Get<WORD>();
			}

			if (resAnd.cRes && resAnd.cRes < _MaxEntriesExtraLarge && ulDepth < _MaxDepth)
			{
				resAnd.lpRes.reserve(resAnd.cRes);
				for (ULONG i = 0; i < resAnd.cRes; i++)
				{
					if (!parser->RemainingBytes()) break;
					resAnd.lpRes.emplace_back(parser, ulDepth + 1, bRuleCondition, bExtendedCount);
				}

				// TODO: Should we do this?
				//srRestriction.resAnd.lpRes.shrink_to_fit();
			}
			break;
		case RES_OR:
			if (!bRuleCondition || bExtendedCount)
			{
				resOr.cRes = parser->Get<DWORD>();
			}
			else
			{
				resOr.cRes = parser->Get<WORD>();
			}

			if (resOr.cRes && resOr.cRes < _MaxEntriesExtraLarge && ulDepth < _MaxDepth)
			{
				resOr.lpRes.reserve(resOr.cRes);
				for (ULONG i = 0; i < resOr.cRes; i++)
				{
					if (!parser->RemainingBytes()) break;
					resOr.lpRes.emplace_back(parser, ulDepth + 1, bRuleCondition, bExtendedCount);
				}
			}
			break;
		case RES_NOT:
			if (ulDepth < _MaxDepth && parser->RemainingBytes())
			{
				resNot.lpRes.emplace_back(parser, ulDepth + 1, bRuleCondition, bExtendedCount);
			}
			break;
		case RES_CONTENT:
			resContent.ulFuzzyLevel = parser->Get<DWORD>();
			resContent.ulPropTag = parser->Get<DWORD>();
			resContent.lpProp.init(parser, 1, bRuleCondition);
			break;
		case RES_PROPERTY:
			if (bRuleCondition)
				resProperty.relop = parser->Get<BYTE>();
			else
				resProperty.relop = parser->Get<DWORD>();

			resProperty.ulPropTag = parser->Get<DWORD>();
			resProperty.lpProp.init(parser, 1, bRuleCondition);
			break;
		case RES_COMPAREPROPS:
			if (bRuleCondition)
				resCompareProps.relop = parser->Get<BYTE>();
			else
				resCompareProps.relop = parser->Get<DWORD>();

			resCompareProps.ulPropTag1 = parser->Get<DWORD>();
			resCompareProps.ulPropTag2 = parser->Get<DWORD>();
			break;
		case RES_BITMASK:
			if (bRuleCondition)
				resBitMask.relBMR = parser->Get<BYTE>();
			else
				resBitMask.relBMR = parser->Get<DWORD>();

			resBitMask.ulPropTag = parser->Get<DWORD>();
			resBitMask.ulMask = parser->Get<DWORD>();
			break;
		case RES_SIZE:
			if (bRuleCondition)
				resSize.relop = parser->Get<BYTE>();
			else
				resSize.relop = parser->Get<DWORD>();

			resSize.ulPropTag = parser->Get<DWORD>();
			resSize.cb = parser->Get<DWORD>();
			break;
		case RES_EXIST:
			resExist.ulPropTag = parser->Get<DWORD>();
			break;
		case RES_SUBRESTRICTION:
			resSub.ulSubObject = parser->Get<DWORD>();
			if (ulDepth < _MaxDepth && parser->RemainingBytes())
			{
				resSub.lpRes.emplace_back(parser, ulDepth + 1, bRuleCondition, bExtendedCount);
			}
			break;
		case RES_COMMENT:
			if (bRuleCondition)
				resComment.cValues = parser->Get<BYTE>();
			else
				resComment.cValues = parser->Get<DWORD>();

			resComment.lpProp.init(parser, resComment.cValues, bRuleCondition);

			// Check if a restriction is present
			if (parser->Get<BYTE>() && ulDepth < _MaxDepth && parser->RemainingBytes())
			{
				resComment.lpRes.emplace_back(parser, ulDepth + 1, bRuleCondition, bExtendedCount);
			}
			break;
		case RES_COUNT:
			resCount.ulCount = parser->Get<DWORD>();
			if (ulDepth < _MaxDepth && parser->RemainingBytes())
			{
				resCount.lpRes.emplace_back(parser, ulDepth + 1, bRuleCondition, bExtendedCount);
			}
			break;
		}
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
		auto szFlags = flags::InterpretFlags(flagRestrictionType, lpRes.rt);
		addBlock(
			lpRes.rt, L"%1!ws!lpRes->rt = 0x%2!X! = %3!ws!\r\n", szTabs.c_str(), lpRes.rt.getData(), szFlags.c_str());

		auto i = 0;
		switch (lpRes.rt)
		{
		case RES_COMPAREPROPS:
			szFlags = flags::InterpretFlags(flagRelop, lpRes.resCompareProps.relop);
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
				proptags::TagToString(lpRes.resCompareProps.ulPropTag1, nullptr, false, true).c_str());
			addBlock(
				lpRes.resCompareProps.ulPropTag2,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag2 = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(lpRes.resCompareProps.ulPropTag2, nullptr, false, true).c_str());
			break;
		case RES_AND:
			addBlock(
				lpRes.resAnd.cRes,
				L"%1!ws!lpRes->res.resAnd.cRes = 0x%2!08X!\r\n",
				szTabs.c_str(),
				lpRes.resAnd.cRes.getData());
			for (const auto& res : lpRes.resAnd.lpRes)
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
			for (const auto& res : lpRes.resOr.lpRes)
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
			if (!lpRes.resNot.lpRes.empty())
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

			if (!lpRes.resCount.lpRes.empty())
			{
				ParseRestriction(lpRes.resCount.lpRes[0], ulTabLevel + 1);
			}
			break;
		case RES_CONTENT:
			szFlags = flags::InterpretFlags(flagFuzzyLevel, lpRes.resContent.ulFuzzyLevel);
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
				proptags::TagToString(lpRes.resContent.ulPropTag, nullptr, false, true).c_str());

			if (!lpRes.resContent.lpProp.Props().empty())
			{
				addBlock(
					lpRes.resContent.lpProp.Props()[0].ulPropTag,
					L"%1!ws!lpRes->res.resContent.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					proptags::TagToString(lpRes.resContent.lpProp.Props()[0].ulPropTag, nullptr, false, true).c_str());
				addBlock(
					lpRes.resContent.lpProp.Props()[0].PropBlock(),
					L"%1!ws!lpRes->res.resContent.lpProp->Value = %2!ws!\r\n",
					szTabs.c_str(),
					lpRes.resContent.lpProp.Props()[0].PropBlock().c_str());
				addBlock(
					lpRes.resContent.lpProp.Props()[0].AltPropBlock(),
					L"%1!ws!\tAlt: %2!ws!\r\n",
					szTabs.c_str(),
					lpRes.resContent.lpProp.Props()[0].AltPropBlock().c_str());
			}
			break;
		case RES_PROPERTY:
			szFlags = flags::InterpretFlags(flagRelop, lpRes.resProperty.relop);
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
				proptags::TagToString(lpRes.resProperty.ulPropTag, nullptr, false, true).c_str());
			if (!lpRes.resProperty.lpProp.Props().empty())
			{
				addBlock(
					lpRes.resProperty.lpProp.Props()[0].ulPropTag,
					L"%1!ws!lpRes->res.resProperty.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					proptags::TagToString(lpRes.resProperty.lpProp.Props()[0].ulPropTag, nullptr, false, true).c_str());
				addBlock(
					lpRes.resProperty.lpProp.Props()[0].PropBlock(),
					L"%1!ws!lpRes->res.resProperty.lpProp->Value = %2!ws!\r\n",
					szTabs.c_str(),
					lpRes.resProperty.lpProp.Props()[0].PropBlock().c_str());
				addBlock(
					lpRes.resProperty.lpProp.Props()[0].AltPropBlock(),
					L"%1!ws!\tAlt: %2!ws!\r\n",
					szTabs.c_str(),
					lpRes.resProperty.lpProp.Props()[0].AltPropBlock().c_str());
				szPropNum = lpRes.resProperty.lpProp.Props()[0].PropNum();
				if (!szPropNum.empty())
				{
					// TODO: use block
					addHeader(L"%1!ws!\tFlags: %2!ws!", szTabs.c_str(), szPropNum.c_str());
				}
			}
			break;
		case RES_BITMASK:
			szFlags = flags::InterpretFlags(flagBitmask, lpRes.resBitMask.relBMR);
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

			terminateBlock();
			addBlock(
				lpRes.resBitMask.ulPropTag,
				L"%1!ws!lpRes->res.resBitMask.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(lpRes.resBitMask.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_SIZE:
			szFlags = flags::InterpretFlags(flagRelop, lpRes.resSize.relop);
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
				proptags::TagToString(lpRes.resSize.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_EXIST:
			addBlock(
				lpRes.resExist.ulPropTag,
				L"%1!ws!lpRes->res.resExist.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(lpRes.resExist.ulPropTag, nullptr, false, true).c_str());
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
				proptags::TagToString(lpRes.resSub.ulSubObject, nullptr, false, true).c_str());
			addHeader(L"%1!ws!lpRes->res.resSub.lpRes\r\n", szTabs.c_str());

			if (!lpRes.resSub.lpRes.empty())
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
					proptags::TagToString(prop.ulPropTag, nullptr, false, true).c_str());
				addBlock(
					prop.PropBlock(),
					L"%1!ws!lpRes->res.resComment.lpProp[0x%2!08X!].Value = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					prop.PropBlock().c_str());
				addBlock(prop.AltPropBlock(), L"%1!ws!\tAlt: %2!ws!\r\n", szTabs.c_str(), prop.AltPropBlock().c_str());
				i++;
			}

			addHeader(L"%1!ws!lpRes->res.resComment.lpRes\r\n", szTabs.c_str());
			if (!lpRes.resComment.lpRes.empty())
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
					proptags::TagToString(prop.ulPropTag, nullptr, false, true).c_str());
				addBlock(
					prop.PropBlock(),
					L"%1!ws!lpRes->res.resAnnotation.lpProp[0x%2!08X!].Value = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					prop.PropBlock().c_str());
				addBlock(prop.AltPropBlock(), L"%1!ws!\tAlt: %2!ws!\r\n", szTabs.c_str(), prop.AltPropBlock().c_str());
				i++;
			}

			addHeader(L"%1!ws!lpRes->res.resAnnotation.lpRes\r\n", szTabs.c_str());

			if (!lpRes.resComment.lpRes.empty())
			{
				ParseRestriction(lpRes.resComment.lpRes[0], ulTabLevel + 1);
			}
			break;
		}
	}
} // namespace smartview