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

	// 2.2.5.1.2.1 OP_MOVE and OP_COPY ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/78bc6a96-f94c-43c6-afd3-7f8c39a2340c

	// 2.2.5.1.2.1.1 ServerEid Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/07e8c314-0ab2-440e-9138-b96f93682bf1

	// 2.2.5.1.2.2 OP_REPLY and OP_OOF_REPLY ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/d5c14d2f-3557-4b78-acfe-d949ef325792

	// 2.2.5.1.2.3 OP_DEFER_ACTION ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/1d66ff31-fdc1-4b36-a040-75207e31dd18

	// 2.2.5.1.2.4 OP_FORWARD and OP_DELEGATE ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/1b8a2b3f-9fff-4e07-9820-c3d2287a8e5c

	// 2.2.5.1.2.4.1 RecipientBlockData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/c6ad133d-7906-43aa-8420-3b40ac6be494

	// 2.2.5.1.2.5 OP_BOUNCE ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/c6ceb0c2-96a2-4337-a4fb-f2e23cfe4284

	// 2.2.5.1.2.6 OP_TAG ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/2fe6b110-5e1e-4a97-aeeb-9103cf76a0e0

	// 2.2.5.1.2.7 OP_DELETE or OP_MARK_AS_READ ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/c180c7cb-474a-46f2-95ee-568ec61f1043

	class RuleAction : public smartViewParser
	{
	public:
		void Init(bool bExtended) noexcept;

	private:
		void parse() override;
		void parseBlocks() override;

		bool m_bExtended{};
	};
} // namespace smartview