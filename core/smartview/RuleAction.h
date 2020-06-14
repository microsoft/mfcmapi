#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXORULE] 2.2.1.3.1.10 PidTagRuleActions Property
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/27dc69c6-c28e-4dd2-84a8-7015b7cded44

	// [MS-OXORULE] 2.2.4.1.9 PidTagExtendedRuleMessageActions Property
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/86d88af3-8c76-4737-aa51-6a299964ba01

	// 2.2.5.1.1 Action Flavors
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/07f635f5-c521-41a5-b01e-712a9ba781d2

	// [MS-OXORULE] 2.2.5.1.2 ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/0c74ec4f-896d-425e-9e6d-20c531fe83ce
	class ActionData : public block
	{
	public:
		ActionData() = default;
		void parse(std::shared_ptr<binaryParser>& _parser, bool bExtended)
		{
			parser = _parser;
			m_bExtended = bExtended;
			ensureParsed();
		}

	protected:
		void parse() override {}
		virtual void parseBlocks() = 0;
		bool m_bExtended{};
	};

	// [MS-OXORULE] 2.2.5.1 ActionBlock Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/0c74ec4f-896d-425e-9e6d-20c531fe83ce
	class ActionBlock : public block
	{
	public:
		ActionBlock(bool bExtended) : m_bExtended(bExtended) {}

	private:
		bool m_bExtended{};
		std::shared_ptr<blockT<DWORD>> ActionLength = emptyT<DWORD>();
		std::shared_ptr<blockT<BYTE>> ActionType = emptyT<BYTE>();
		std::shared_ptr<blockT<DWORD>> ActionFlavor = emptyT<DWORD>();
		std::shared_ptr<ActionData> ActionData;

		void parse() override;
		void parseBlocks() override;
	};

	// [MS-OXORULE] 2.2.5 RuleAction Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/a31c67b8-405e-4c4f-bfda-e5be0091041a
	class RuleAction : public block
	{
	public:
		void Init(bool bExtended) noexcept { m_bExtended = bExtended; }

	private:
		bool m_bExtended{};
		std::shared_ptr<blockT<DWORD>> NoOfActions = emptyT<DWORD>();
		std::vector<std::shared_ptr<ActionBlock>> ActionBlocks;
		void parse() override
		{
			NoOfActions = m_bExtended ? blockT<DWORD>::parse(parser) : blockT<DWORD>::parse<WORD>(parser);
			if (*NoOfActions < _MaxEntriesSmall)
			{
				ActionBlocks.reserve(*NoOfActions);
				for (auto i = 0; i < *NoOfActions; i++)
				{
					auto actionBlock = std::make_shared<ActionBlock>(m_bExtended);
					actionBlock->block::parse(parser, false);
					ActionBlocks.push_back(actionBlock);
				}
			}
		}

		void parseBlocks() override
		{
			setText(m_bExtended ? L"Extended Rule Action\r\n" : L"Rule Action\r\n");
			addChild(NoOfActions);
			for (const auto actionBlock : ActionBlocks)
			{
				addChild(actionBlock);
				terminateBlock();
			}
		}
	};
} // namespace smartview