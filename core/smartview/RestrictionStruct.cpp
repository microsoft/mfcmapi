#include <core/stdafx.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	// If bRuleCondition is true, parse restrictions as defined in [MS-OXCDATA] 2.12
	// If bRuleCondition is true, bExtendedCount controls whether the count fields in AND/OR restrictions is 16 or 32 bits
	//   https://msdn.microsoft.com/en-us/library/ee201126(v=exchg.80).aspx
	// If bRuleCondition is false, parse restrictions as defined in [MS-OXOCFG] 2.2.6.1.2
	// If bRuleCondition is false, ignore bExtendedCount (assumes true)
	//   https://msdn.microsoft.com/en-us/library/ee217813(v=exchg.80).aspx
	// Never fails, but will not parse restrictions above _MaxDepth
	// [MS-OXCDATA] 2.11.4 TaggedPropertyValue Structure
	void RestrictionStruct::parse(ULONG ulDepth)
	{
		if (m_bRuleCondition)
		{
			rt = m_Parser->Get<BYTE>();
		}
		else
		{
			rt = m_Parser->Get<DWORD>();
		}

		switch (rt)
		{
		case RES_AND:
			if (!m_bRuleCondition || m_bExtendedCount)
			{
				resAnd.cRes = m_Parser->Get<DWORD>();
			}
			else
			{
				resAnd.cRes = m_Parser->Get<WORD>();
			}

			if (resAnd.cRes && resAnd.cRes < _MaxEntriesExtraLarge && ulDepth < _MaxDepth)
			{
				resAnd.lpRes.reserve(resAnd.cRes);
				for (ULONG i = 0; i < resAnd.cRes; i++)
				{
					if (!m_Parser->RemainingBytes()) break;
					resAnd.lpRes.emplace_back(
						std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount));
				}

				// TODO: Should we do this?
				//srRestriction.resAnd.lpRes.shrink_to_fit();
			}
			break;
		case RES_OR:
			if (!m_bRuleCondition || m_bExtendedCount)
			{
				resOr.cRes = m_Parser->Get<DWORD>();
			}
			else
			{
				resOr.cRes = m_Parser->Get<WORD>();
			}

			if (resOr.cRes && resOr.cRes < _MaxEntriesExtraLarge && ulDepth < _MaxDepth)
			{
				resOr.lpRes.reserve(resOr.cRes);
				for (ULONG i = 0; i < resOr.cRes; i++)
				{
					if (!m_Parser->RemainingBytes()) break;
					resOr.lpRes.emplace_back(
						std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount));
				}
			}
			break;
		case RES_NOT:
			if (ulDepth < _MaxDepth && m_Parser->RemainingBytes())
			{
				resNot.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
			break;
		case RES_CONTENT:
			resContent.ulFuzzyLevel = m_Parser->Get<DWORD>();
			resContent.ulPropTag = m_Parser->Get<DWORD>();
			resContent.lpProp.init(m_Parser, 1, m_bRuleCondition);
			break;
		case RES_PROPERTY:
			if (m_bRuleCondition)
				resProperty.relop = m_Parser->Get<BYTE>();
			else
				resProperty.relop = m_Parser->Get<DWORD>();

			resProperty.ulPropTag = m_Parser->Get<DWORD>();
			resProperty.lpProp.init(m_Parser, 1, m_bRuleCondition);
			break;
		case RES_COMPAREPROPS:
			if (m_bRuleCondition)
				resCompareProps.relop = m_Parser->Get<BYTE>();
			else
				resCompareProps.relop = m_Parser->Get<DWORD>();

			resCompareProps.ulPropTag1 = m_Parser->Get<DWORD>();
			resCompareProps.ulPropTag2 = m_Parser->Get<DWORD>();
			break;
		case RES_BITMASK:
			if (m_bRuleCondition)
				resBitMask.relBMR = m_Parser->Get<BYTE>();
			else
				resBitMask.relBMR = m_Parser->Get<DWORD>();

			resBitMask.ulPropTag = m_Parser->Get<DWORD>();
			resBitMask.ulMask = m_Parser->Get<DWORD>();
			break;
		case RES_SIZE:
			if (m_bRuleCondition)
				resSize.relop = m_Parser->Get<BYTE>();
			else
				resSize.relop = m_Parser->Get<DWORD>();

			resSize.ulPropTag = m_Parser->Get<DWORD>();
			resSize.cb = m_Parser->Get<DWORD>();
			break;
		case RES_EXIST:
			resExist.ulPropTag = m_Parser->Get<DWORD>();
			break;
		case RES_SUBRESTRICTION:
			resSub.ulSubObject = m_Parser->Get<DWORD>();
			if (ulDepth < _MaxDepth && m_Parser->RemainingBytes())
			{
				resSub.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
			break;
		case RES_COMMENT:
			if (m_bRuleCondition)
				resComment.cValues = m_Parser->Get<BYTE>();
			else
				resComment.cValues = m_Parser->Get<DWORD>();

			resComment.lpProp.init(m_Parser, resComment.cValues, m_bRuleCondition);

			// Check if a restriction is present
			if (m_Parser->Get<BYTE>() && ulDepth < _MaxDepth && m_Parser->RemainingBytes())
			{
				resComment.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}

			break;
		case RES_COUNT:
			resCount.ulCount = m_Parser->Get<DWORD>();
			if (ulDepth < _MaxDepth && m_Parser->RemainingBytes())
			{
				resCount.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}

			break;
		}
	}

	// There may be restrictions with over 100 nested levels, but we're not going to try to parse them
#define _MaxRestrictionNesting 100

	void RestrictionStruct::parseBlocks(ULONG ulTabLevel)
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
		auto szFlags = flags::InterpretFlags(flagRestrictionType, rt);
		addBlock(rt, L"%1!ws!lpRes->rt = 0x%2!X! = %3!ws!\r\n", szTabs.c_str(), rt.getData(), szFlags.c_str());

		auto i = 0;
		switch (rt)
		{
		case RES_COMPAREPROPS:
			szFlags = flags::InterpretFlags(flagRelop, resCompareProps.relop);
			addBlock(
				resCompareProps.relop,
				L"%1!ws!lpRes->res.resCompareProps.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				resCompareProps.relop.getData());
			addBlock(
				resCompareProps.ulPropTag1,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag1 = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resCompareProps.ulPropTag1, nullptr, false, true).c_str());
			addBlock(
				resCompareProps.ulPropTag2,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag2 = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resCompareProps.ulPropTag2, nullptr, false, true).c_str());
			break;
		case RES_AND:
			addBlock(
				resAnd.cRes, L"%1!ws!lpRes->res.resAnd.cRes = 0x%2!08X!\r\n", szTabs.c_str(), resAnd.cRes.getData());
			for (const auto& res : resAnd.lpRes)
			{
				addHeader(L"%1!ws!lpRes->res.resAnd.lpRes[0x%2!08X!]\r\n", szTabs.c_str(), i++);
				res->parseBlocks(ulTabLevel + 1);
				addBlock(res->getBlock());
			}
			break;
		case RES_OR:
			addBlock(resOr.cRes, L"%1!ws!lpRes->res.resOr.cRes = 0x%2!08X!\r\n", szTabs.c_str(), resOr.cRes.getData());
			for (const auto& res : resOr.lpRes)
			{
				addHeader(L"%1!ws!lpRes->res.resOr.lpRes[0x%2!08X!]\r\n", szTabs.c_str(), i++);
				res->parseBlocks(ulTabLevel + 1);
				addBlock(res->getBlock());
			}
			break;
		case RES_NOT:
			addBlock(
				resNot.ulReserved,
				L"%1!ws!lpRes->res.resNot.ulReserved = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resNot.ulReserved.getData());
			addHeader(L"%1!ws!lpRes->res.resNot.lpRes\r\n", szTabs.c_str());

			if (resNot.lpRes)
			{
				resNot.lpRes->parseBlocks(ulTabLevel + 1);
				addBlock(resNot.lpRes->getBlock());
			}
			break;
		case RES_COUNT:
			addBlock(
				resCount.ulCount,
				L"%1!ws!lpRes->res.resCount.ulCount = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resCount.ulCount.getData());
			addHeader(L"%1!ws!lpRes->res.resCount.lpRes\r\n", szTabs.c_str());

			if (resCount.lpRes)
			{
				resCount.lpRes->parseBlocks(ulTabLevel + 1);
				addBlock(resCount.lpRes->getBlock());
			}
			break;
		case RES_CONTENT:
			szFlags = flags::InterpretFlags(flagFuzzyLevel, resContent.ulFuzzyLevel);
			addBlock(
				resContent.ulFuzzyLevel,
				L"%1!ws!lpRes->res.resContent.ulFuzzyLevel = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				resContent.ulFuzzyLevel.getData());
			addBlock(
				resContent.ulPropTag,
				L"%1!ws!lpRes->res.resContent.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resContent.ulPropTag, nullptr, false, true).c_str());

			if (!resContent.lpProp.Props().empty())
			{
				addBlock(
					resContent.lpProp.Props()[0].ulPropTag,
					L"%1!ws!lpRes->res.resContent.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					proptags::TagToString(resContent.lpProp.Props()[0].ulPropTag, nullptr, false, true).c_str());
				addBlock(
					resContent.lpProp.Props()[0].PropBlock(),
					L"%1!ws!lpRes->res.resContent.lpProp->Value = %2!ws!\r\n",
					szTabs.c_str(),
					resContent.lpProp.Props()[0].PropBlock().c_str());
				addBlock(
					resContent.lpProp.Props()[0].AltPropBlock(),
					L"%1!ws!\tAlt: %2!ws!\r\n",
					szTabs.c_str(),
					resContent.lpProp.Props()[0].AltPropBlock().c_str());
			}
			break;
		case RES_PROPERTY:
			szFlags = flags::InterpretFlags(flagRelop, resProperty.relop);
			addBlock(
				resProperty.relop,
				L"%1!ws!lpRes->res.resProperty.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				resProperty.relop.getData());
			addBlock(
				resProperty.ulPropTag,
				L"%1!ws!lpRes->res.resProperty.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resProperty.ulPropTag, nullptr, false, true).c_str());
			if (!resProperty.lpProp.Props().empty())
			{
				addBlock(
					resProperty.lpProp.Props()[0].ulPropTag,
					L"%1!ws!lpRes->res.resProperty.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					proptags::TagToString(resProperty.lpProp.Props()[0].ulPropTag, nullptr, false, true).c_str());
				addBlock(
					resProperty.lpProp.Props()[0].PropBlock(),
					L"%1!ws!lpRes->res.resProperty.lpProp->Value = %2!ws!\r\n",
					szTabs.c_str(),
					resProperty.lpProp.Props()[0].PropBlock().c_str());
				addBlock(
					resProperty.lpProp.Props()[0].AltPropBlock(),
					L"%1!ws!\tAlt: %2!ws!\r\n",
					szTabs.c_str(),
					resProperty.lpProp.Props()[0].AltPropBlock().c_str());
				szPropNum = resProperty.lpProp.Props()[0].PropNum();
				if (!szPropNum.empty())
				{
					// TODO: use block
					addHeader(L"%1!ws!\tFlags: %2!ws!", szTabs.c_str(), szPropNum.c_str());
				}
			}
			break;
		case RES_BITMASK:
			szFlags = flags::InterpretFlags(flagBitmask, resBitMask.relBMR);
			addBlock(
				resBitMask.relBMR,
				L"%1!ws!lpRes->res.resBitMask.relBMR = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				resBitMask.relBMR.getData());
			addBlock(
				resBitMask.ulMask,
				L"%1!ws!lpRes->res.resBitMask.ulMask = 0x%2!08X!",
				szTabs.c_str(),
				resBitMask.ulMask.getData());
			szPropNum = InterpretNumberAsStringProp(resBitMask.ulMask, resBitMask.ulPropTag);
			if (!szPropNum.empty())
			{
				addBlock(resBitMask.ulMask, L": %1!ws!", szPropNum.c_str());
			}

			terminateBlock();
			addBlock(
				resBitMask.ulPropTag,
				L"%1!ws!lpRes->res.resBitMask.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resBitMask.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_SIZE:
			szFlags = flags::InterpretFlags(flagRelop, resSize.relop);
			addBlock(
				resSize.relop,
				L"%1!ws!lpRes->res.resSize.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				szFlags.c_str(),
				resSize.relop.getData());
			addBlock(resSize.cb, L"%1!ws!lpRes->res.resSize.cb = 0x%2!08X!\r\n", szTabs.c_str(), resSize.cb.getData());
			addBlock(
				resSize.ulPropTag,
				L"%1!ws!lpRes->res.resSize.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resSize.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_EXIST:
			addBlock(
				resExist.ulPropTag,
				L"%1!ws!lpRes->res.resExist.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resExist.ulPropTag, nullptr, false, true).c_str());
			addBlock(
				resExist.ulReserved1,
				L"%1!ws!lpRes->res.resExist.ulReserved1 = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resExist.ulReserved1.getData());
			addBlock(
				resExist.ulReserved2,
				L"%1!ws!lpRes->res.resExist.ulReserved2 = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resExist.ulReserved2.getData());
			break;
		case RES_SUBRESTRICTION:
			addBlock(
				resSub.ulSubObject,
				L"%1!ws!lpRes->res.resSub.ulSubObject = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resSub.ulSubObject, nullptr, false, true).c_str());
			addHeader(L"%1!ws!lpRes->res.resSub.lpRes\r\n", szTabs.c_str());

			if (resSub.lpRes)
			{
				resSub.lpRes->parseBlocks(ulTabLevel + 1);
				addBlock(resSub.lpRes->getBlock());
			}
			break;
		case RES_COMMENT:
			addBlock(
				resComment.cValues,
				L"%1!ws!lpRes->res.resComment.cValues = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resComment.cValues.getData());
			for (auto prop : resComment.lpProp.Props())
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
			if (resComment.lpRes)
			{
				resComment.lpRes->parseBlocks(ulTabLevel + 1);
				addBlock(resComment.lpRes->getBlock());
			}
			break;
		case RES_ANNOTATION:
			addBlock(
				resComment.cValues,
				L"%1!ws!lpRes->res.resAnnotation.cValues = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resComment.cValues.getData());
			for (auto prop : resComment.lpProp.Props())
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

			if (resComment.lpRes)
			{
				resComment.lpRes->parseBlocks(ulTabLevel + 1);
				addBlock(resComment.lpRes->getBlock());
			}

			break;
		}
	}
} // namespace smartview