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
			rt.parse<BYTE>(m_Parser);
		}
		else
		{
			rt.parse<DWORD>(m_Parser);
		}

		switch (rt)
		{
		case RES_AND:
			if (!m_bRuleCondition || m_bExtendedCount)
			{
				resAnd.cRes.parse<DWORD>(m_Parser);
			}
			else
			{
				resAnd.cRes.parse<WORD>(m_Parser);
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
				resOr.cRes.parse<DWORD>(m_Parser);
			}
			else
			{
				resOr.cRes.parse<WORD>(m_Parser);
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
			resContent.ulFuzzyLevel.parse<DWORD>(m_Parser);
			resContent.ulPropTag.parse<DWORD>(m_Parser);
			resContent.lpProp.init(m_Parser, 1, m_bRuleCondition);
			break;
		case RES_PROPERTY:
			if (m_bRuleCondition)
				resProperty.relop.parse<BYTE>(m_Parser);
			else
				resProperty.relop.parse<DWORD>(m_Parser);

			resProperty.ulPropTag.parse<DWORD>(m_Parser);
			resProperty.lpProp.init(m_Parser, 1, m_bRuleCondition);
			break;
		case RES_COMPAREPROPS:
			if (m_bRuleCondition)
				resCompareProps.relop.parse<BYTE>(m_Parser);
			else
				resCompareProps.relop.parse<DWORD>(m_Parser);

			resCompareProps.ulPropTag1.parse<DWORD>(m_Parser);
			resCompareProps.ulPropTag2.parse<DWORD>(m_Parser);
			break;
		case RES_BITMASK:
			if (m_bRuleCondition)
				resBitMask.relBMR.parse<BYTE>(m_Parser);
			else
				resBitMask.relBMR.parse<DWORD>(m_Parser);

			resBitMask.ulPropTag.parse<DWORD>(m_Parser);
			resBitMask.ulMask.parse<DWORD>(m_Parser);
			break;
		case RES_SIZE:
			if (m_bRuleCondition)
				resSize.relop.parse<BYTE>(m_Parser);
			else
				resSize.relop.parse<DWORD>(m_Parser);

			resSize.ulPropTag.parse<DWORD>(m_Parser);
			resSize.cb.parse<DWORD>(m_Parser);
			break;
		case RES_EXIST:
			resExist.ulPropTag.parse<DWORD>(m_Parser);
			break;
		case RES_SUBRESTRICTION:
			resSub.ulSubObject.parse<DWORD>(m_Parser);
			if (ulDepth < _MaxDepth && m_Parser->RemainingBytes())
			{
				resSub.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
			break;
		case RES_COMMENT:
		{
			if (m_bRuleCondition)
				resComment.cValues.parse<BYTE>(m_Parser);
			else
				resComment.cValues.parse<DWORD>(m_Parser);

			resComment.lpProp.init(m_Parser, resComment.cValues, m_bRuleCondition);

			// Check if a restriction is present
			const auto& resExists = blockT<BYTE>(m_Parser);
			if (resExists && ulDepth < _MaxDepth && m_Parser->RemainingBytes())
			{
				resComment.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		break;
		case RES_ANNOTATION:
		{
			if (m_bRuleCondition)
				resAnnotation.cValues.parse<BYTE>(m_Parser);
			else
				resAnnotation.cValues.parse<DWORD>(m_Parser);

			resAnnotation.lpProp.init(m_Parser, resAnnotation.cValues, m_bRuleCondition);

			// Check if a restriction is present
			const auto& resExists = blockT<BYTE>(m_Parser);
			if (resExists && ulDepth < _MaxDepth && m_Parser->RemainingBytes())
			{
				resAnnotation.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		break;
		case RES_COUNT:
			resCount.ulCount.parse<DWORD>(m_Parser);
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

		auto i = 0;
		switch (rt)
		{
		case RES_COMPAREPROPS:
			rt.addBlock(
				resCompareProps.relop,
				L"%1!ws!lpRes->res.resCompareProps.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagRelop, resCompareProps.relop).c_str(),
				resCompareProps.relop.getData());
			rt.addBlock(
				resCompareProps.ulPropTag1,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag1 = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resCompareProps.ulPropTag1, nullptr, false, true).c_str());
			rt.addBlock(
				resCompareProps.ulPropTag2,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag2 = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resCompareProps.ulPropTag2, nullptr, false, true).c_str());
			break;
		case RES_AND:
			for (const auto& res : resAnd.lpRes)
			{
				auto resBlock = std::make_shared<block>();
				resBlock->setText(L"%1!ws!lpRes->res.resAnd.lpRes[0x%2!08X!]\r\n", szTabs.c_str(), i++);
				res->parseBlocks(ulTabLevel + 1);
				resBlock->addBlock(res->getBlock());
				resAnd.cRes.addBlock(resBlock);
			}

			rt.addBlock(
				resAnd.cRes, L"%1!ws!lpRes->res.resAnd.cRes = 0x%2!08X!\r\n", szTabs.c_str(), resAnd.cRes.getData());
			break;
		case RES_OR:
			for (const auto& res : resOr.lpRes)
			{
				auto resBlock = block{};
				resBlock.setText(L"%1!ws!lpRes->res.resOr.lpRes[0x%2!08X!]\r\n", szTabs.c_str(), i++);
				res->parseBlocks(ulTabLevel + 1);
				resBlock.addBlock(res->getBlock());
				resOr.cRes.addBlock(resBlock);
			}

			rt.addBlock(
				resOr.cRes, L"%1!ws!lpRes->res.resOr.cRes = 0x%2!08X!\r\n", szTabs.c_str(), resOr.cRes.getData());
			break;
		case RES_NOT:
		{
			rt.addBlock(
				resNot.ulReserved,
				L"%1!ws!lpRes->res.resNot.ulReserved = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resNot.ulReserved.getData());

			auto notRoot = block{};
			notRoot.setText(L"%1!ws!lpRes->res.resNot.lpRes\r\n", szTabs.c_str());

			if (resNot.lpRes)
			{
				resNot.lpRes->parseBlocks(ulTabLevel + 1);
				notRoot.addBlock(resNot.lpRes->getBlock());
			}

			rt.addBlock(notRoot);
		}

		break;
		case RES_COUNT:
		{
			rt.addBlock(
				resCount.ulCount,
				L"%1!ws!lpRes->res.resCount.ulCount = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resCount.ulCount.getData());

			auto countRoot = block{};
			countRoot.setText(L"%1!ws!lpRes->res.resCount.lpRes\r\n", szTabs.c_str());

			if (resCount.lpRes)
			{
				resCount.lpRes->parseBlocks(ulTabLevel + 1);
				countRoot.addBlock(resCount.lpRes->getBlock());
			}

			rt.addBlock(countRoot);
		}

		break;
		case RES_CONTENT:
		{
			rt.addBlock(
				resContent.ulFuzzyLevel,
				L"%1!ws!lpRes->res.resContent.ulFuzzyLevel = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagFuzzyLevel, resContent.ulFuzzyLevel).c_str(),
				resContent.ulFuzzyLevel.getData());

			rt.addBlock(
				resContent.ulPropTag,
				L"%1!ws!lpRes->res.resContent.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resContent.ulPropTag, nullptr, false, true).c_str());

			if (!resContent.lpProp.Props().empty())
			{
				auto propBlock = block{};
				propBlock.addBlock(
					resContent.lpProp.Props()[0]->ulPropTag,
					L"%1!ws!lpRes->res.resContent.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					proptags::TagToString(resContent.lpProp.Props()[0]->ulPropTag, nullptr, false, true).c_str());
				propBlock.addBlock(
					resContent.lpProp.Props()[0]->PropBlock(),
					L"%1!ws!lpRes->res.resContent.lpProp->Value = %2!ws!\r\n",
					szTabs.c_str(),
					resContent.lpProp.Props()[0]->PropBlock().c_str());
				propBlock.addBlock(
					resContent.lpProp.Props()[0]->AltPropBlock(),
					L"%1!ws!\tAlt: %2!ws!\r\n",
					szTabs.c_str(),
					resContent.lpProp.Props()[0]->AltPropBlock().c_str());
				rt.addBlock(propBlock);
			}
		}
		break;
		case RES_PROPERTY:
		{
			rt.addBlock(
				resProperty.relop,
				L"%1!ws!lpRes->res.resProperty.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagRelop, resProperty.relop).c_str(),
				resProperty.relop.getData());
			if (!resProperty.lpProp.Props().empty())
			{
				resProperty.ulPropTag.addBlock(
					resProperty.lpProp.Props()[0]->ulPropTag,
					L"%1!ws!lpRes->res.resProperty.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					proptags::TagToString(resProperty.lpProp.Props()[0]->ulPropTag, nullptr, false, true).c_str());
				resProperty.ulPropTag.addBlock(
					resProperty.lpProp.Props()[0]->PropBlock(),
					L"%1!ws!lpRes->res.resProperty.lpProp->Value = %2!ws!\r\n",
					szTabs.c_str(),
					resProperty.lpProp.Props()[0]->PropBlock().c_str());
				resProperty.ulPropTag.addBlock(
					resProperty.lpProp.Props()[0]->AltPropBlock(),
					L"%1!ws!\tAlt: %2!ws!\r\n",
					szTabs.c_str(),
					resProperty.lpProp.Props()[0]->AltPropBlock().c_str());
				szPropNum = resProperty.lpProp.Props()[0]->PropNum();
				if (!szPropNum.empty())
				{
					// TODO: use block
					resProperty.ulPropTag.addHeader(L"%1!ws!\tFlags: %2!ws!", szTabs.c_str(), szPropNum.c_str());
				}
			}

			rt.addBlock(
				resProperty.ulPropTag,
				L"%1!ws!lpRes->res.resProperty.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resProperty.ulPropTag, nullptr, false, true).c_str());
		}

		break;
		case RES_BITMASK:
			rt.addBlock(
				resBitMask.relBMR,
				L"%1!ws!lpRes->res.resBitMask.relBMR = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagBitmask, resBitMask.relBMR).c_str(),
				resBitMask.relBMR.getData());
			rt.addBlock(
				resBitMask.ulMask,
				L"%1!ws!lpRes->res.resBitMask.ulMask = 0x%2!08X!",
				szTabs.c_str(),
				resBitMask.ulMask.getData());
			szPropNum = InterpretNumberAsStringProp(resBitMask.ulMask, resBitMask.ulPropTag);
			if (!szPropNum.empty())
			{
				rt.addBlock(resBitMask.ulMask, L": %1!ws!", szPropNum.c_str());
			}

			rt.terminateBlock();
			rt.addBlock(
				resBitMask.ulPropTag,
				L"%1!ws!lpRes->res.resBitMask.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resBitMask.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_SIZE:
			rt.addBlock(
				resSize.relop,
				L"%1!ws!lpRes->res.resSize.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagRelop, resSize.relop).c_str(),
				resSize.relop.getData());
			rt.addBlock(
				resSize.cb, L"%1!ws!lpRes->res.resSize.cb = 0x%2!08X!\r\n", szTabs.c_str(), resSize.cb.getData());
			rt.addBlock(
				resSize.ulPropTag,
				L"%1!ws!lpRes->res.resSize.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resSize.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_EXIST:
			rt.addBlock(
				resExist.ulPropTag,
				L"%1!ws!lpRes->res.resExist.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resExist.ulPropTag, nullptr, false, true).c_str());
			rt.addBlock(
				resExist.ulReserved1,
				L"%1!ws!lpRes->res.resExist.ulReserved1 = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resExist.ulReserved1.getData());
			rt.addBlock(
				resExist.ulReserved2,
				L"%1!ws!lpRes->res.resExist.ulReserved2 = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resExist.ulReserved2.getData());
			break;
		case RES_SUBRESTRICTION:
		{
			resSub.ulSubObject.addHeader(L"%1!ws!lpRes->res.resSub.lpRes\r\n", szTabs.c_str());

			if (resSub.lpRes)
			{
				resSub.lpRes->parseBlocks(ulTabLevel + 1);
				resSub.ulSubObject.addBlock(resSub.lpRes->getBlock());
			}

			rt.addBlock(
				resSub.ulSubObject,
				L"%1!ws!lpRes->res.resSub.ulSubObject = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(resSub.ulSubObject, nullptr, false, true).c_str());
		}

		break;
		case RES_COMMENT:
		{
			for (const auto& prop : resComment.lpProp.Props())
			{
				prop->ulPropTag.setText(
					L"%1!ws!lpRes->res.resComment.lpProp[0x%2!08X!].ulPropTag = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					proptags::TagToString(prop->ulPropTag, nullptr, false, true).c_str());
				prop->ulPropTag.addBlock(
					prop->PropBlock(),
					L"%1!ws!lpRes->res.resComment.lpProp[0x%2!08X!].Value = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					prop->PropBlock().c_str());
				prop->ulPropTag.addBlock(
					prop->AltPropBlock(), L"%1!ws!\tAlt: %2!ws!\r\n", szTabs.c_str(), prop->AltPropBlock().c_str());

				resComment.cValues.addBlock(prop->ulPropTag);
				i++;
			}

			rt.addBlock(
				resComment.cValues,
				L"%1!ws!lpRes->res.resComment.cValues = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resComment.cValues.getData());

			rt.addHeader(L"%1!ws!lpRes->res.resComment.lpRes\r\n", szTabs.c_str());
			if (resComment.lpRes)
			{
				resComment.lpRes->parseBlocks(ulTabLevel + 1);
				rt.addBlock(resComment.lpRes->getBlock());
			}
		}

		break;
		case RES_ANNOTATION:
			for (const auto& prop : resAnnotation.lpProp.Props())
			{
				prop->ulPropTag.setText(
					L"%1!ws!lpRes->res.resAnnotation.lpProp[0x%2!08X!].ulPropTag = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					proptags::TagToString(prop->ulPropTag, nullptr, false, true).c_str());
				prop->ulPropTag.addBlock(
					prop->PropBlock(),
					L"%1!ws!lpRes->res.resAnnotation.lpProp[0x%2!08X!].Value = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					prop->PropBlock().c_str());
				prop->ulPropTag.addBlock(
					prop->AltPropBlock(), L"%1!ws!\tAlt: %2!ws!\r\n", szTabs.c_str(), prop->AltPropBlock().c_str());

				resAnnotation.cValues.addBlock(prop->ulPropTag);
				i++;
			}

			rt.addBlock(
				resAnnotation.cValues,
				L"%1!ws!lpRes->res.resAnnotation.cValues = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resAnnotation.cValues.getData());

			rt.addHeader(L"%1!ws!lpRes->res.resAnnotation.lpRes\r\n", szTabs.c_str());
			if (resAnnotation.lpRes)
			{
				resAnnotation.lpRes->parseBlocks(ulTabLevel + 1);
				rt.addBlock(resAnnotation.lpRes->getBlock());
			}

			break;
		}

		addBlock(
			rt,
			L"%1!ws!lpRes->rt = 0x%2!X! = %3!ws!\r\n",
			szTabs.c_str(),
			rt.getData(),
			flags::InterpretFlags(flagRestrictionType, rt).c_str());
	} // namespace smartview
} // namespace smartview