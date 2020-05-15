#include <core/stdafx.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/3d0e5e14-7464-4659-b83d-a601b81fbdf5
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/001660fe-9839-41b9-9d5f-cd2b1a577e3b
	class SAndRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (!m_bRuleCondition || m_bExtendedCount)
			{
				cRes = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				cRes = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (*cRes && *cRes < _MaxEntriesExtraLarge && m_ulDepth < _MaxDepth)
			{
				lpRes.reserve(*cRes);
				for (ULONG i = 0; i < *cRes; i++)
				{
					if (!m_Parser->getSize()) break;
					lpRes.emplace_back(std::make_shared<RestrictionStruct>(
						m_Parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount));
				}

				// TODO: Should we do this?
				//srRestriction.resAnd.lpRes.shrink_to_fit();
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			const auto t = makeTabs(ulTabLevel);
			const auto tabs = t.c_str();

			auto i = 0;
			addChild(cRes, L"%1!ws!lpRes->res.resAnd.cRes = 0x%2!08X!\r\n", tabs, cRes->getData());

			for (const auto& res : lpRes)
			{
				auto resBlock = std::make_shared<block>();
				resBlock->setText(L"%1!ws!lpRes->res.resAnd.lpRes[0x%2!08X!]\r\n", tabs, i++);
				res->parseBlocks(ulTabLevel + 1);
				resBlock->addChild(res->getBlock());
				cRes->addChild(resBlock);
			}
		}

		std::shared_ptr<blockT<DWORD>> cRes = emptyT<DWORD>();
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/50a969db-a794-4e3c-9fc0-28f37bc1d7a2
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/09499990-ab00-4d4f-9c0d-cff61d9cddff
	class SOrRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (!m_bRuleCondition || m_bExtendedCount)
			{
				cRes = blockT<DWORD>::parse(m_Parser);
			}
			else
			{
				cRes = blockT<DWORD, WORD>::parse(m_Parser);
			}

			if (*cRes && *cRes < _MaxEntriesExtraLarge && m_ulDepth < _MaxDepth)
			{
				lpRes.reserve(*cRes);
				for (ULONG i = 0; i < *cRes; i++)
				{
					if (!m_Parser->getSize()) break;
					lpRes.emplace_back(std::make_shared<RestrictionStruct>(
						m_Parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount));
				}
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			const auto t = makeTabs(ulTabLevel);
			const auto tabs = t.c_str();

			auto i = 0;
			addChild(cRes, L"%1!ws!lpRes->res.resOr.cRes = 0x%2!08X!\r\n", tabs, cRes->getData());

			for (const auto& res : lpRes)
			{
				auto resBlock = std::make_shared<block>();
				resBlock->setText(L"%1!ws!lpRes->res.resOr.lpRes[0x%2!08X!]\r\n", tabs, i++);
				res->parseBlocks(ulTabLevel + 1);
				resBlock->addChild(res->getBlock());
				cRes->addChild(resBlock);
			}
		}

		std::shared_ptr<blockT<DWORD>> cRes = emptyT<DWORD>();
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};
	// If bRuleCondition is true, parse restrictions as defined in [MS-OXCDATA] 2.12
	// If bRuleCondition is true, bExtendedCount controls whether the count fields in AND/OR restrictions is 16 or 32 bits
	//   https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/5d554ba7-b82f-42b6-8802-97c19f760633
	// If bRuleCondition is false, parse restrictions as defined in [MS-OXOCFG] 2.2.6.1.2
	// If bRuleCondition is false, ignore bExtendedCount (assumes true)
	//   https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/6d9acc02-dafa-494b-88ae-1c4385e547d6
	// Never fails, but will not parse restrictions above _MaxDepth
	// [MS-OXCDATA] 2.11.4 TaggedPropertyValue Structure
	void RestrictionStruct::parse(ULONG ulDepth)
	{
		if (m_bRuleCondition)
		{
			rt = blockT<DWORD, BYTE>::parse(m_Parser);
		}
		else
		{
			rt = blockT<DWORD>::parse(m_Parser);
		}

		switch (*rt)
		{
		case RES_AND:
			res1 = std::make_shared<SAndRestrictionStruct>();
			res1->parse(m_Parser, ulDepth, m_bRuleCondition, m_bExtendedCount);
			break;
		case RES_OR:
			res1 = std::make_shared<SOrRestrictionStruct>();
			res1->parse(m_Parser, ulDepth, m_bRuleCondition, m_bExtendedCount);
			break;
		case RES_NOT:
			if (ulDepth < _MaxDepth && m_Parser->getSize())
			{
				resNot.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
			break;
		case RES_CONTENT:
			resContent.ulFuzzyLevel = blockT<DWORD>::parse(m_Parser);
			resContent.ulPropTag = blockT<DWORD>::parse(m_Parser);
			resContent.lpProp.parse(m_Parser, 1, m_bRuleCondition);
			break;
		case RES_PROPERTY:
			if (m_bRuleCondition)
				resProperty.relop = blockT<DWORD, BYTE>::parse(m_Parser);
			else
				resProperty.relop = blockT<DWORD>::parse(m_Parser);

			resProperty.ulPropTag = blockT<DWORD>::parse(m_Parser);
			resProperty.lpProp.parse(m_Parser, 1, m_bRuleCondition);
			break;
		case RES_COMPAREPROPS:
			if (m_bRuleCondition)
				resCompareProps.relop = blockT<DWORD, BYTE>::parse(m_Parser);
			else
				resCompareProps.relop = blockT<DWORD>::parse(m_Parser);

			resCompareProps.ulPropTag1 = blockT<DWORD>::parse(m_Parser);
			resCompareProps.ulPropTag2 = blockT<DWORD>::parse(m_Parser);
			break;
		case RES_BITMASK:
			if (m_bRuleCondition)
				resBitMask.relBMR = blockT<DWORD, BYTE>::parse(m_Parser);
			else
				resBitMask.relBMR = blockT<DWORD>::parse(m_Parser);

			resBitMask.ulPropTag = blockT<DWORD>::parse(m_Parser);
			resBitMask.ulMask = blockT<DWORD>::parse(m_Parser);
			break;
		case RES_SIZE:
			if (m_bRuleCondition)
				resSize.relop = blockT<DWORD, BYTE>::parse(m_Parser);
			else
				resSize.relop = blockT<DWORD>::parse(m_Parser);

			resSize.ulPropTag = blockT<DWORD>::parse(m_Parser);
			resSize.cb = blockT<DWORD>::parse(m_Parser);
			break;
		case RES_EXIST:
			resExist.ulPropTag = blockT<DWORD>::parse(m_Parser);
			break;
		case RES_SUBRESTRICTION:
			resSub.ulSubObject = blockT<DWORD>::parse(m_Parser);
			if (ulDepth < _MaxDepth && m_Parser->getSize())
			{
				resSub.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
			break;
		case RES_COMMENT:
		{
			if (m_bRuleCondition)
				resComment.cValues = blockT<DWORD, BYTE>::parse(m_Parser);
			else
				resComment.cValues = blockT<DWORD>::parse(m_Parser);

			resComment.lpProp.parse(m_Parser, *resComment.cValues, m_bRuleCondition);

			// Check if a restriction is present
			const auto resExists = blockT<BYTE>::parse(m_Parser);
			if (*resExists && ulDepth < _MaxDepth && m_Parser->getSize())
			{
				resComment.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		break;
		case RES_ANNOTATION:
		{
			if (m_bRuleCondition)
				resAnnotation.cValues = blockT<DWORD, BYTE>::parse(m_Parser);
			else
				resAnnotation.cValues = blockT<DWORD>::parse(m_Parser);

			resAnnotation.lpProp.parse(m_Parser, *resAnnotation.cValues, m_bRuleCondition);

			// Check if a restriction is present
			const auto& resExists = blockT<BYTE>::parse(m_Parser);
			if (*resExists && ulDepth < _MaxDepth && m_Parser->getSize())
			{
				resAnnotation.lpRes =
					std::make_shared<RestrictionStruct>(m_Parser, ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		break;
		case RES_COUNT:
			resCount.ulCount = blockT<DWORD>::parse(m_Parser);
			if (ulDepth < _MaxDepth && m_Parser->getSize())
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

		addChild(
			rt,
			L"%1!ws!lpRes->rt = 0x%2!X! = %3!ws!\r\n",
			szTabs.c_str(),
			rt->getData(),
			flags::InterpretFlags(flagRestrictionType, *rt).c_str());

		switch (*rt)
		{
		case RES_COMPAREPROPS:
			rt->addChild(
				resCompareProps.relop,
				L"%1!ws!lpRes->res.resCompareProps.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagRelop, *resCompareProps.relop).c_str(),
				resCompareProps.relop->getData());
			rt->addChild(
				resCompareProps.ulPropTag1,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag1 = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(*resCompareProps.ulPropTag1, nullptr, false, true).c_str());
			rt->addChild(
				resCompareProps.ulPropTag2,
				L"%1!ws!lpRes->res.resCompareProps.ulPropTag2 = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(*resCompareProps.ulPropTag2, nullptr, false, true).c_str());
			break;
		case RES_AND:
		case RES_OR:
			if (res1)
			{
				res1->parseBlocks(ulTabLevel);
				addChild(res1);
			}

			break;
		case RES_NOT:
		{
			auto notRoot = std::make_shared<block>();
			notRoot->setText(L"%1!ws!lpRes->res.resNot.lpRes\r\n", szTabs.c_str());
			rt->addChild(notRoot);

			if (resNot.lpRes)
			{
				resNot.lpRes->parseBlocks(ulTabLevel + 1);
				notRoot->addChild(resNot.lpRes->getBlock());
			}
		}

		break;
		case RES_COUNT:
		{
			rt->addChild(
				resCount.ulCount,
				L"%1!ws!lpRes->res.resCount.ulCount = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resCount.ulCount->getData());

			auto countRoot = std::make_shared<block>();
			countRoot->setText(L"%1!ws!lpRes->res.resCount.lpRes\r\n", szTabs.c_str());
			rt->addChild(countRoot);

			if (resCount.lpRes)
			{
				resCount.lpRes->parseBlocks(ulTabLevel + 1);
				countRoot->addChild(resCount.lpRes->getBlock());
			}
		}

		break;
		case RES_CONTENT:
		{
			rt->addChild(
				resContent.ulFuzzyLevel,
				L"%1!ws!lpRes->res.resContent.ulFuzzyLevel = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagFuzzyLevel, *resContent.ulFuzzyLevel).c_str(),
				resContent.ulFuzzyLevel->getData());

			rt->addChild(
				resContent.ulPropTag,
				L"%1!ws!lpRes->res.resContent.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(*resContent.ulPropTag, nullptr, false, true).c_str());

			if (!resContent.lpProp.Props().empty())
			{
				auto propBlock = std::make_shared<block>();
				rt->addChild(propBlock);

				propBlock->addChild(
					resContent.lpProp.Props()[0]->ulPropTag,
					L"%1!ws!lpRes->res.resContent.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					proptags::TagToString(*resContent.lpProp.Props()[0]->ulPropTag, nullptr, false, true).c_str());
				if (resContent.lpProp.Props()[0]->value)
				{
					if (!resContent.lpProp.Props()[0]->value->PropBlock()->empty())
					{
						propBlock->addChild(
							resContent.lpProp.Props()[0]->value->PropBlock(),
							L"%1!ws!lpRes->res.resContent.lpProp->Value = %2!ws!\r\n",
							szTabs.c_str(),
							resContent.lpProp.Props()[0]->value->PropBlock()->c_str());
					}

					if (!resContent.lpProp.Props()[0]->value->AltPropBlock()->empty())
					{
						propBlock->addChild(
							resContent.lpProp.Props()[0]->value->AltPropBlock(),
							L"%1!ws!\tAlt: %2!ws!\r\n",
							szTabs.c_str(),
							resContent.lpProp.Props()[0]->value->AltPropBlock()->c_str());
					}
				}
			}
		}
		break;
		case RES_PROPERTY:
		{
			rt->addChild(
				resProperty.relop,
				L"%1!ws!lpRes->res.resProperty.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagRelop, *resProperty.relop).c_str(),
				resProperty.relop->getData());
			rt->addChild(
				resProperty.ulPropTag,
				L"%1!ws!lpRes->res.resProperty.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(*resProperty.ulPropTag, nullptr, false, true).c_str());

			if (!resProperty.lpProp.Props().empty())
			{
				resProperty.ulPropTag->addChild(
					resProperty.lpProp.Props()[0]->ulPropTag,
					L"%1!ws!lpRes->res.resProperty.lpProp->ulPropTag = %2!ws!\r\n",
					szTabs.c_str(),
					proptags::TagToString(*resProperty.lpProp.Props()[0]->ulPropTag, nullptr, false, true).c_str());
				if (resProperty.lpProp.Props()[0]->value)
				{
					if (!resProperty.lpProp.Props()[0]->value->PropBlock()->empty())
					{
						resProperty.ulPropTag->addChild(
							resProperty.lpProp.Props()[0]->value->PropBlock(),
							L"%1!ws!lpRes->res.resProperty.lpProp->Value = %2!ws!\r\n",
							szTabs.c_str(),
							resProperty.lpProp.Props()[0]->value->PropBlock()->c_str());
					}

					if (!resProperty.lpProp.Props()[0]->value->AltPropBlock()->empty())
					{
						resProperty.ulPropTag->addChild(
							resProperty.lpProp.Props()[0]->value->AltPropBlock(),
							L"%1!ws!\tAlt: %2!ws!\r\n",
							szTabs.c_str(),
							resProperty.lpProp.Props()[0]->value->AltPropBlock()->c_str());
					}

					szPropNum = resProperty.lpProp.Props()[0]->value->toNumberAsString();
					if (!szPropNum.empty())
					{
						// TODO: use block
						resProperty.ulPropTag->addHeader(L"%1!ws!\tFlags: %2!ws!", szTabs.c_str(), szPropNum.c_str());
					}
				}
			}
		}

		break;
		case RES_BITMASK:
			rt->addChild(
				resBitMask.relBMR,
				L"%1!ws!lpRes->res.resBitMask.relBMR = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagBitmask, *resBitMask.relBMR).c_str(),
				resBitMask.relBMR->getData());
			rt->addChild(
				resBitMask.ulPropTag,
				L"%1!ws!lpRes->res.resBitMask.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(*resBitMask.ulPropTag, nullptr, false, true).c_str());
			szPropNum = InterpretNumberAsStringProp(*resBitMask.ulMask, *resBitMask.ulPropTag);
			if (szPropNum.empty())
			{
				rt->addChild(
					resBitMask.ulMask,
					L"%1!ws!lpRes->res.resBitMask.ulMask = 0x%2!08X!\r\n",
					szTabs.c_str(),
					resBitMask.ulMask->getData());
			}
			else
			{
				rt->addChild(
					resBitMask.ulMask,
					L"%1!ws!lpRes->res.resBitMask.ulMask = %2!ws! = 0x%3!08X!\r\n",
					szTabs.c_str(),
					szPropNum.c_str(),
					resBitMask.ulMask->getData());
			}

			break;
		case RES_SIZE:
			rt->addChild(
				resSize.relop,
				L"%1!ws!lpRes->res.resSize.relop = %2!ws! = 0x%3!08X!\r\n",
				szTabs.c_str(),
				flags::InterpretFlags(flagRelop, *resSize.relop).c_str(),
				resSize.relop->getData());
			rt->addChild(
				resSize.cb, L"%1!ws!lpRes->res.resSize.cb = 0x%2!08X!\r\n", szTabs.c_str(), resSize.cb->getData());
			rt->addChild(
				resSize.ulPropTag,
				L"%1!ws!lpRes->res.resSize.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(*resSize.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_EXIST:
			rt->addChild(
				resExist.ulPropTag,
				L"%1!ws!lpRes->res.resExist.ulPropTag = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(*resExist.ulPropTag, nullptr, false, true).c_str());
			break;
		case RES_SUBRESTRICTION:
		{
			rt->addChild(
				resSub.ulSubObject,
				L"%1!ws!lpRes->res.resSub.ulSubObject = %2!ws!\r\n",
				szTabs.c_str(),
				proptags::TagToString(*resSub.ulSubObject, nullptr, false, true).c_str());
			resSub.ulSubObject->addHeader(L"%1!ws!lpRes->res.resSub.lpRes\r\n", szTabs.c_str());

			if (resSub.lpRes)
			{
				resSub.lpRes->parseBlocks(ulTabLevel + 1);
				resSub.ulSubObject->addChild(resSub.lpRes->getBlock());
			}
		}

		break;
		case RES_COMMENT:
		{
			rt->addChild(
				resComment.cValues,
				L"%1!ws!lpRes->res.resComment.cValues = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resComment.cValues->getData());

			auto i = 0;
			for (const auto& prop : resComment.lpProp.Props())
			{
				prop->ulPropTag->setText(
					L"%1!ws!lpRes->res.resComment.lpProp[0x%2!08X!].ulPropTag = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					proptags::TagToString(*prop->ulPropTag, nullptr, false, true).c_str());
				resComment.cValues->addChild(prop->ulPropTag);

				if (prop->value)
				{
					if (!prop->value->PropBlock()->empty())
					{
						prop->ulPropTag->addChild(
							prop->value->PropBlock(),
							L"%1!ws!lpRes->res.resComment.lpProp[0x%2!08X!].Value = %3!ws!\r\n",
							szTabs.c_str(),
							i,
							prop->value->PropBlock()->c_str());
					}

					if (!prop->value->AltPropBlock()->empty())
					{
						prop->ulPropTag->addChild(
							prop->value->AltPropBlock(),
							L"%1!ws!\tAlt: %2!ws!\r\n",
							szTabs.c_str(),
							prop->value->AltPropBlock()->c_str());
					}
				}

				i++;
			}

			rt->addHeader(L"%1!ws!lpRes->res.resComment.lpRes\r\n", szTabs.c_str());
			if (resComment.lpRes)
			{
				resComment.lpRes->parseBlocks(ulTabLevel + 1);
				rt->addChild(resComment.lpRes->getBlock());
			}
		}

		break;
		case RES_ANNOTATION:
		{
			rt->addChild(
				resAnnotation.cValues,
				L"%1!ws!lpRes->res.resAnnotation.cValues = 0x%2!08X!\r\n",
				szTabs.c_str(),
				resAnnotation.cValues->getData());

			auto i = 0;
			for (const auto& prop : resAnnotation.lpProp.Props())
			{
				resAnnotation.cValues->addChild(prop->ulPropTag);

				prop->ulPropTag->setText(
					L"%1!ws!lpRes->res.resAnnotation.lpProp[0x%2!08X!].ulPropTag = %3!ws!\r\n",
					szTabs.c_str(),
					i,
					proptags::TagToString(*prop->ulPropTag, nullptr, false, true).c_str());
				if (prop->value)
				{
					if (!prop->value->PropBlock()->empty())
					{
						prop->ulPropTag->addChild(
							prop->value->PropBlock(),
							L"%1!ws!lpRes->res.resAnnotation.lpProp[0x%2!08X!].Value = %3!ws!\r\n",
							szTabs.c_str(),
							i,
							prop->value->PropBlock()->c_str());
					}

					if (!prop->value->AltPropBlock()->empty())
					{
						prop->ulPropTag->addChild(
							prop->value->AltPropBlock(),
							L"%1!ws!\tAlt: %2!ws!\r\n",
							szTabs.c_str(),
							prop->value->AltPropBlock()->c_str());
					}
				}

				i++;
			}

			rt->addHeader(L"%1!ws!lpRes->res.resAnnotation.lpRes\r\n", szTabs.c_str());
			if (resAnnotation.lpRes)
			{
				resAnnotation.lpRes->parseBlocks(ulTabLevel + 1);
				rt->addChild(resAnnotation.lpRes->getBlock());
			}
		}
		break;
		}
	} // namespace smartview
} // namespace smartview