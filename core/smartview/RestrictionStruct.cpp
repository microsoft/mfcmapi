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
				cRes = blockT<DWORD>::parse(parser);
			}
			else
			{
				cRes = blockT<DWORD>::parse<WORD>(parser);
			}

			if (*cRes && *cRes < _MaxEntriesExtraLarge && m_ulDepth < _MaxDepth)
			{
				lpRes.reserve(*cRes);
				for (ULONG i = 0; i < *cRes; i++)
				{
					if (!parser->getSize()) break;
					lpRes.emplace_back(
						std::make_shared<RestrictionStruct>(parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount));
				}

				// TODO: Should we do this?
				//srRestriction.resAnd.lpRes.shrink_to_fit();
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resAnd");
			addChild(cRes, L"cRes = 0x%1!08X!", cRes->getData());

			auto i = 0;
			for (const auto& res : lpRes)
			{
				res->parseBlocks(ulTabLevel + 1);
				addChild(res, L"lpRes[0x%1!08X!]", i++);
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
				cRes = blockT<DWORD>::parse(parser);
			}
			else
			{
				cRes = blockT<DWORD>::parse<WORD>(parser);
			}

			if (*cRes && *cRes < _MaxEntriesExtraLarge && m_ulDepth < _MaxDepth)
			{
				lpRes.reserve(*cRes);
				for (ULONG i = 0; i < *cRes; i++)
				{
					if (!parser->getSize()) break;
					lpRes.emplace_back(
						std::make_shared<RestrictionStruct>(parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount));
				}
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resOr");
			addChild(cRes, L"cRes = 0x%1!08X!", cRes->getData());

			auto i = 0;
			for (const auto& res : lpRes)
			{
				res->parseBlocks(ulTabLevel + 1);
				addChild(res, L"lpRes[0x%1!08X!]", i++);
			}
		}

		std::shared_ptr<blockT<DWORD>> cRes = emptyT<DWORD>();
		std::vector<std::shared_ptr<RestrictionStruct>> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/0abc5c41-9db7-4e6c-8d4d-b5c7e51d5355
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/f10fbc18-b384-4cd8-9490-a6955035e2ec
	class SNotRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (m_ulDepth < _MaxDepth && parser->getSize())
			{
				lpRes = std::make_shared<RestrictionStruct>(parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resNot");

			if (lpRes)
			{
				lpRes->parseBlocks(ulTabLevel + 1);
				addChild(lpRes, L"lpRes");
			}
		}

		std::shared_ptr<RestrictionStruct> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/486b4b66-a6f0-43d3-98ee-91f1e0e29907
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/e0dc196d-e041-4d80-bc44-89b94fd71158
	class SComparePropsRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (m_bRuleCondition)
				relop = blockT<DWORD>::parse<BYTE>(parser);
			else
				relop = blockT<DWORD>::parse(parser);

			ulPropTag1 = blockT<DWORD>::parse(parser);
			ulPropTag2 = blockT<DWORD>::parse(parser);
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resCompareProps");
			addChild(
				relop,
				L"relop = %1!ws! = 0x%2!08X!",
				flags::InterpretFlags(flagRelop, *relop).c_str(),
				relop->getData());
			addChild(
				ulPropTag1, L"ulPropTag1 = %1!ws!", proptags::TagToString(*ulPropTag1, nullptr, false, true).c_str());
			addChild(
				ulPropTag2, L"ulPropTag2 = %1!ws!", proptags::TagToString(*ulPropTag2, nullptr, false, true).c_str());
		}

		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag1 = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag2 = emptyT<DWORD>();
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/c1526deb-d05d-4d42-af68-d0233e4cd064
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/e6ded216-f49e-49b9-8b26-e1743e76c897
	class SContentRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			ulFuzzyLevel = blockT<DWORD>::parse(parser);
			ulPropTag = blockT<DWORD>::parse(parser);
			lpProp.parse(parser, 1, m_bRuleCondition);
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resContent");
			addChild(
				ulFuzzyLevel,
				L"ulFuzzyLevel = %1!ws! = 0x%2!08X!",
				flags::InterpretFlags(flagFuzzyLevel, *ulFuzzyLevel).c_str(),
				ulFuzzyLevel->getData());

			addChild(ulPropTag, L"ulPropTag = %1!ws!", proptags::TagToString(*ulPropTag, nullptr, false, true).c_str());

			if (!lpProp.Props().empty())
			{
				auto propBlock = create(L"lpProp");
				addChild(propBlock);

				propBlock->addChild(
					lpProp.Props()[0]->ulPropTag,
					L"ulPropTag = %1!ws!",
					proptags::TagToString(*lpProp.Props()[0]->ulPropTag, nullptr, false, true).c_str());
				if (lpProp.Props()[0]->value)
				{
					if (!lpProp.Props()[0]->value->PropBlock()->empty())
					{
						propBlock->addChild(
							lpProp.Props()[0]->value->PropBlock(),
							L"Value = %1!ws!",
							lpProp.Props()[0]->value->PropBlock()->c_str());
					}

					if (!lpProp.Props()[0]->value->AltPropBlock()->empty())
					{
						propBlock->addChild(
							lpProp.Props()[0]->value->AltPropBlock(),
							L"Alt: %1!ws!",
							lpProp.Props()[0]->value->AltPropBlock()->c_str());
					}
				}
			}
		}

		std::shared_ptr<blockT<DWORD>> ulFuzzyLevel = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		PropertiesStruct lpProp;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/66b892da-032a-4f66-812a-01c11457d5ce
	class SCountRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			ulCount = blockT<DWORD>::parse(parser);
			if (m_ulDepth < _MaxDepth && parser->getSize())
			{
				lpRes = std::make_shared<RestrictionStruct>(parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resCount");
			addChild(ulCount, L"resCount.ulCount = 0x%1!08X!", ulCount->getData());

			if (lpRes)
			{
				auto lpResRoot = create(L"lpRes");
				addChild(lpResRoot);
				lpRes->parseBlocks(ulTabLevel + 1);
				addChild(lpRes);
			}
		}

		std::shared_ptr<blockT<DWORD>> ulCount = emptyT<DWORD>();
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/285d931b-e1ef-4390-aa84-9a3ee98b6e28
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/4e66da53-032b-46f0-8b84-39ee1e96de80
	class SPropertyRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (m_bRuleCondition)
				relop = blockT<DWORD>::parse<BYTE>(parser);
			else
				relop = blockT<DWORD>::parse(parser);

			ulPropTag = blockT<DWORD>::parse(parser);
			lpProp.parse(parser, 1, m_bRuleCondition);
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resProperty");
			addChild(
				relop,
				L"relop = %1!ws! = 0x%2!08X!",
				flags::InterpretFlags(flagRelop, *relop).c_str(),
				relop->getData());
			addChild(ulPropTag, L"ulPropTag = %1!ws!", proptags::TagToString(*ulPropTag, nullptr, false, true).c_str());

			if (!lpProp.Props().empty())
			{
				auto propBlock = create(L"lpProp");
				addChild(propBlock);
				propBlock->addChild(
					lpProp.Props()[0]->ulPropTag,
					L"ulPropTag = %1!ws!",
					proptags::TagToString(*lpProp.Props()[0]->ulPropTag, nullptr, false, true).c_str());
				if (lpProp.Props()[0]->value)
				{
					if (!lpProp.Props()[0]->value->PropBlock()->empty())
					{
						propBlock->addChild(
							lpProp.Props()[0]->value->PropBlock(),
							L"Value = %1!ws!",
							lpProp.Props()[0]->value->PropBlock()->c_str());
					}

					if (!lpProp.Props()[0]->value->AltPropBlock()->empty())
					{
						propBlock->addChild(
							lpProp.Props()[0]->value->AltPropBlock(),
							L"Alt: %1!ws!",
							lpProp.Props()[0]->value->AltPropBlock()->c_str());
					}

					const auto szPropNum = lpProp.Props()[0]->value->toNumberAsString();
					if (!szPropNum.empty())
					{
						// TODO: use block
						propBlock->addHeader(L"Flags: %1!ws!", szPropNum.c_str());
					}
				}
			}
		}

		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		PropertiesStruct lpProp;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/c1526deb-d05d-4d42-af68-d0233e4cd064
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/e6ded216-f49e-49b9-8b26-e1743e76c897
	class SBitMaskRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (m_bRuleCondition)
				relBMR = blockT<DWORD>::parse<BYTE>(parser);
			else
				relBMR = blockT<DWORD>::parse(parser);

			ulPropTag = blockT<DWORD>::parse(parser);
			ulMask = blockT<DWORD>::parse(parser);
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resBitMask");
			addChild(
				relBMR,
				L"relBMR = %1!ws! = 0x%2!08X!",
				flags::InterpretFlags(flagBitmask, *relBMR).c_str(),
				relBMR->getData());
			addChild(ulPropTag, L"ulPropTag = %1!ws!", proptags::TagToString(*ulPropTag, nullptr, false, true).c_str());
			const auto szPropNum = InterpretNumberAsStringProp(*ulMask, *ulPropTag);
			if (szPropNum.empty())
			{
				addChild(ulMask, L"ulMask = 0x%1!08X!", ulMask->getData());
			}
			else
			{
				addChild(ulMask, L"ulMask = %1!ws! = 0x%2!08X!", szPropNum.c_str(), ulMask->getData());
			}
		}

		std::shared_ptr<blockT<DWORD>> relBMR = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulMask = emptyT<DWORD>();
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/b81c6f60-0e17-43e0-8d1f-c3015c7cd9de
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/c6915f61-debe-466c-9145-ffadee0cb28e
	class SSizeRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (m_bRuleCondition)
				relop = blockT<DWORD>::parse<BYTE>(parser);
			else
				relop = blockT<DWORD>::parse(parser);

			ulPropTag = blockT<DWORD>::parse(parser);
			cb = blockT<DWORD>::parse(parser);
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resSize");
			addChild(
				relop,
				L"relop = %1!ws! = 0x%2!08X!",
				flags::InterpretFlags(flagRelop, *relop).c_str(),
				relop->getData());
			addChild(cb, L"cb = 0x%1!08X!", cb->getData());
			addChild(ulPropTag, L"ulPropTag = %1!ws!", proptags::TagToString(*ulPropTag, nullptr, false, true).c_str());
		}

		std::shared_ptr<blockT<DWORD>> relop = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> cb = emptyT<DWORD>();
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/c4da9872-8903-4b70-9b23-9980f6bc21c7
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/116661e8-6c17-417c-a75d-261ac0ac29ba
	class SExistRestrictionStruct : public blockRes
	{
	public:
		void parse() override { ulPropTag = blockT<DWORD>::parse(parser); }

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resExist");
			addChild(ulPropTag, L"ulPropTag = %1!ws!", proptags::TagToString(*ulPropTag, nullptr, false, true).c_str());
		}

		std::shared_ptr<blockT<DWORD>> ulPropTag = emptyT<DWORD>();
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/9e87191a-fe53-48a1-aba4-55ba1f8c221f
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/550dc0ac-01cf-4ee2-b15c-799faa852120
	class SSubRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			ulSubObject = blockT<DWORD>::parse(parser);
			if (m_ulDepth < _MaxDepth && parser->getSize())
			{
				lpRes = std::make_shared<RestrictionStruct>(parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resSub");
			addChild(
				ulSubObject,
				L"ulSubObject = %1!ws!",
				proptags::TagToString(*ulSubObject, nullptr, false, true).c_str());

			if (lpRes)
			{
				lpRes->parseBlocks(ulTabLevel + 1);
				addChild(lpRes, L"lpRes");
			}
		}

		std::shared_ptr<blockT<DWORD>> ulSubObject = emptyT<DWORD>();
		std::shared_ptr<RestrictionStruct> lpRes;
	};

	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/ceb8bd59-90e5-4c6f-8632-8fc9ef639766
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/2ffa8c7a-38c1-4446-8fe1-13b52854460e
	class SCommentRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (m_bRuleCondition)
				cValues = blockT<DWORD>::parse<BYTE>(parser);
			else
				cValues = blockT<DWORD>::parse(parser);

			lpProp.parse(parser, *cValues, m_bRuleCondition);

			// Check if a restriction is present
			const auto resExists = blockT<BYTE>::parse(parser);
			if (*resExists && m_ulDepth < _MaxDepth && parser->getSize())
			{
				lpRes = std::make_shared<RestrictionStruct>(parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resComment");
			addChild(cValues, L"cValues = 0x%1!08X!", cValues->getData());

			auto i = 0;
			for (const auto& prop : lpProp.Props())
			{
				auto propBlock = create(L"lpProp[0x%1!08X!]", i);
				addChild(propBlock);
				propBlock->addChild(
					prop->ulPropTag,
					L"ulPropTag = %1!ws!",
					proptags::TagToString(*prop->ulPropTag, nullptr, false, true).c_str());

				if (prop->value)
				{
					if (!prop->value->PropBlock()->empty())
					{
						propBlock->addChild(
							prop->value->PropBlock(), L"Value = %1!ws!", prop->value->PropBlock()->c_str());
					}

					if (!prop->value->AltPropBlock()->empty())
					{
						propBlock->addChild(
							prop->value->AltPropBlock(), L"Alt: %1!ws!", prop->value->AltPropBlock()->c_str());
					}
				}

				i++;
			}

			if (lpRes)
			{
				lpRes->parseBlocks(ulTabLevel + 1);
				addChild(lpRes, L"lpRes");
			}
		}

		std::shared_ptr<blockT<DWORD>> cValues = emptyT<DWORD>(); /* # of properties in lpProp */
		std::shared_ptr<RestrictionStruct> lpRes;
		PropertiesStruct lpProp;
	};

	class SAnnotationRestrictionStruct : public blockRes
	{
	public:
		void parse() override
		{
			if (m_bRuleCondition)
				cValues = blockT<DWORD>::parse<BYTE>(parser);
			else
				cValues = blockT<DWORD>::parse(parser);

			lpProp.parse(parser, *cValues, m_bRuleCondition);

			// Check if a restriction is present
			const auto& resExists = blockT<BYTE>::parse(parser);
			if (*resExists && m_ulDepth < _MaxDepth && parser->getSize())
			{
				lpRes = std::make_shared<RestrictionStruct>(parser, m_ulDepth + 1, m_bRuleCondition, m_bExtendedCount);
			}
		}

		void parseBlocks(ULONG ulTabLevel)
		{
			setText(L"resAnnotation");
			addChild(cValues, L"cValues = 0x%1!08X!", cValues->getData());

			auto i = 0;
			for (const auto& prop : lpProp.Props())
			{
				auto propBlock = create(L"lpProp[0x%1!08X!]", i);
				addChild(propBlock);
				propBlock->addChild(prop->ulPropTag);

				propBlock->setText(
					L"ulPropTag = %1!ws!", proptags::TagToString(*prop->ulPropTag, nullptr, false, true).c_str());
				if (prop->value)
				{
					if (!prop->value->PropBlock()->empty())
					{
						propBlock->addChild(
							prop->value->PropBlock(), L"Value = %1!ws!", prop->value->PropBlock()->c_str());
					}

					if (!prop->value->AltPropBlock()->empty())
					{
						propBlock->addChild(
							prop->value->AltPropBlock(), L"Alt: %1!ws!", prop->value->AltPropBlock()->c_str());
					}
				}

				i++;
			}

			if (lpRes)
			{
				auto resBlock = create(L"lpRes");
				addChild(resBlock);
				lpRes->parseBlocks(ulTabLevel + 1);
				resBlock->addChild(lpRes);
			}
		}

		std::shared_ptr<blockT<DWORD>> cValues = emptyT<DWORD>(); /* # of properties in lpProp */
		std::shared_ptr<RestrictionStruct> lpRes;
		PropertiesStruct lpProp;
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
			rt = blockT<DWORD>::parse<BYTE>(parser);
		}
		else
		{
			rt = blockT<DWORD>::parse(parser);
		}

		switch (*rt)
		{
		case RES_AND:
			res = std::make_shared<SAndRestrictionStruct>();
			break;
		case RES_OR:
			res = std::make_shared<SOrRestrictionStruct>();
			break;
		case RES_NOT:
			res = std::make_shared<SNotRestrictionStruct>();
			break;
		case RES_CONTENT:
			res = std::make_shared<SContentRestrictionStruct>();
			break;
		case RES_PROPERTY:
			res = std::make_shared<SPropertyRestrictionStruct>();
			break;
		case RES_COMPAREPROPS:
			res = std::make_shared<SComparePropsRestrictionStruct>();
			break;
		case RES_BITMASK:
			res = std::make_shared<SBitMaskRestrictionStruct>();
			break;
		case RES_SIZE:
			res = std::make_shared<SSizeRestrictionStruct>();
			break;
		case RES_EXIST:
			res = std::make_shared<SExistRestrictionStruct>();
			break;
		case RES_SUBRESTRICTION:
			res = std::make_shared<SSubRestrictionStruct>();
			break;
		case RES_COMMENT:
			res = std::make_shared<SCommentRestrictionStruct>();
			break;
		case RES_ANNOTATION:
			res = std::make_shared<SAnnotationRestrictionStruct>();
			break;
		case RES_COUNT:
			res = std::make_shared<SCountRestrictionStruct>();
			break;
		}

		if (res)
		{
			res->parse(parser, ulDepth, m_bRuleCondition, m_bExtendedCount);
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

		addChild(rt, L"rt = 0x%1!X! = %2!ws!", rt->getData(), flags::InterpretFlags(flagRestrictionType, *rt).c_str());

		if (res)
		{
			res->parseBlocks(ulTabLevel);
			addChild(res);
		}
	}
} // namespace smartview