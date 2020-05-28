#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXORULE] 2.2.1.3.1.10 PidTagRuleActions Property
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/27dc69c6-c28e-4dd2-84a8-7015b7cded44

	// [MS-OXORULE] 2.2.4.1.9 PidTagExtendedRuleMessageActions Property
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/86d88af3-8c76-4737-aa51-6a299964ba01

	// [MS-OXORULE] 2.2.5 RuleAction Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/a31c67b8-405e-4c4f-bfda-e5be0091041a

	// [MS-OXORULE] 2.2.5.1 ActionBlock Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/0c74ec4f-896d-425e-9e6d-20c531fe83ce

	// 2.2.5.1.1 Action Flavors
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/07f635f5-c521-41a5-b01e-712a9ba781d2

	// [MS-OXORULE] 2.2.5.1.2 ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/0c74ec4f-896d-425e-9e6d-20c531fe83ce
	class ActionData : public block
	{
	public:
		ActionData() = default;
		void parse(std::shared_ptr<binaryParser>& parser, bool bExtended)
		{
			m_Parser = parser;
			m_bExtended = bExtended;

			// Offset will always be where we start parsing
			setOffset(m_Parser->getOffset());
			parse();
			// And size will always be how many bytes we consumed
			setSize(m_Parser->getOffset() - getOffset());
		}
		ActionData(const ActionData&) = delete;
		ActionData& operator=(const ActionData&) = delete;
		virtual void parseBlocks(ULONG ulTabLevel) = 0;

	protected:
		bool m_bExtended{};
		std::shared_ptr<binaryParser> m_Parser{};

	private:
		virtual void parse() = 0;
	};

	class ActionBlock : public block
	{
	public:
		void parse(std::shared_ptr<binaryParser>& parser, bool bExtended)
		{
			// Offset will always be where we start parsing
			setOffset(parser->getOffset());
			// And size will always be how many bytes we consumed
			setSize(parser->getOffset() - getOffset());
		}

	private:
		bool m_bExtended{};
		std::shared_ptr<blockT<WORD>> ActionLength = emptyT<WORD>();
		std::shared_ptr<blockT<BYTE>> ActionType = emptyT<BYTE>();
		std::shared_ptr<blockT<DWORD>> ActionFlavor = emptyT<DWORD>();
		std::shared_ptr<ActionData> ActionData;
	};

	class RuleAction : public smartViewParser
	{
	public:
		void Init(bool bExtended) noexcept;

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_bExtended{};
		std::shared_ptr<blockT<DWORD>> NoOfActions = emptyT<DWORD>();
		std::vector<std::shared_ptr<ActionBlock>> ActionBlocks;
	};
} // namespace smartview