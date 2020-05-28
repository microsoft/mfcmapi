#include <core/stdafx.h>
#include <core/smartview/RuleAction.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/interpret/guid.h>

namespace smartview
{
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
	class ActionDataBounce : public ActionData
	{
	public:
		void parse() override { BounceCode = blockT<DWORD>::parse(m_Parser); }

		void parseBlocks()
		{
			// TODO: write this
		}

		std::shared_ptr<blockT<DWORD>> BounceCode = emptyT<DWORD>();
	};

	// 2.2.5.1.2.6 OP_TAG ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/2fe6b110-5e1e-4a97-aeeb-9103cf76a0e0
	// AKA TaggedPropertyValue

	// 2.2.5.1.2.7 OP_DELETE or OP_MARK_AS_READ ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/c180c7cb-474a-46f2-95ee-568ec61f1043
	// No structure

	void RuleAction::Init(bool bExtended) noexcept { m_bExtended = bExtended; }

	void RuleAction::parse() {}

	void RuleAction::parseBlocks()
	{
		setRoot(m_bExtended ? L"Extended Rule Action\r\n" : L"Rule Action\r\n");
		NoOfActions = m_bExtended ? blockT<DWORD>::parse(m_Parser) : blockT<DWORD, WORD>::parse(m_Parser);
		ActionBlocks.reserve(*NoOfActions);
		if (*NoOfActions < _MaxEntriesSmall)
		{
			for (auto i = 0; i < *NoOfActions; i++)
			{
				auto actionBlock = std::make_shared<ActionBlock>();
				actionBlock->parse(m_Parser, m_bExtended);
				ActionBlocks.push_back(actionBlock);
			}
		}
	}
} // namespace smartview