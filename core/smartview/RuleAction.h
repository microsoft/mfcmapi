#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>
#include <core/smartview/RuleCondition.h>

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
		void init(BYTE _actionType, bool bExtended)
		{
			actionType = _actionType;
			m_bExtended = bExtended;
		}

	protected:
		void parse() override {}
		virtual void parseBlocks() = 0;
		bool m_bExtended{};
		BYTE actionType;
	};

	// [MS-OXORULE] 2.2.5.1 ActionBlock Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/0c74ec4f-896d-425e-9e6d-20c531fe83ce
	class ActionBlock : public block
	{
	public:
		ActionBlock(bool bExtended) : m_bExtended(bExtended) {}

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_bExtended{};
		std::shared_ptr<blockT<DWORD>> ActionLength = emptyT<DWORD>();
		std::shared_ptr<blockT<BYTE>> ActionType = emptyT<BYTE>();
		std::shared_ptr<blockT<DWORD>> ActionFlavor = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ActionFlags = emptyT<DWORD>();
		std::shared_ptr<ActionData> ActionData;
	};

	// [MS-OXORULE] 2.2.5 RuleAction Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/a31c67b8-405e-4c4f-bfda-e5be0091041a
	class RuleAction : public block
	{
	public:
		RuleAction(bool bExtended) : m_bExtended(bExtended) {}

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_bExtended{};
		std::shared_ptr<NamedPropertyInformation> namedPropertyInformation;
		std::shared_ptr<blockT<DWORD>> RuleVersion = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> NoOfActions = emptyT<DWORD>();
		std::vector<std::shared_ptr<ActionBlock>> ActionBlocks;
	};
} // namespace smartview