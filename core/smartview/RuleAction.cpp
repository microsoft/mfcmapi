#include <core/stdafx.h>
#include <core/smartview/RuleAction.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/interpret/guid.h>

namespace smartview
{
	class ServerEID : public block
	{
	private:
		std::shared_ptr<blockT<BYTE>> Ours = emptyT<BYTE>();
		std::shared_ptr<blockT<LARGE_INTEGER>> FolderId = emptyT<LARGE_INTEGER>();
		std::shared_ptr<blockT<LARGE_INTEGER>> MessageId = emptyT<LARGE_INTEGER>();
		std::shared_ptr<blockT<DWORD>> Instance = emptyT<DWORD>();

		void parse() override
		{
			Ours = blockT<BYTE>::parse(parser);
			FolderId = blockT<LARGE_INTEGER>::parse(parser);
			MessageId = blockT<LARGE_INTEGER>::parse(parser);
			Instance = blockT<DWORD>::parse(parser);
		};
		void parseBlocks() override
		{
			setText(L"ServerEID:\r\n");
			addChild(Ours);
			addChild(FolderId);
			addChild(MessageId);
			addChild(Instance);
		};
	};

	// 2.2.5.1.2.1 OP_MOVE and OP_COPY ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/78bc6a96-f94c-43c6-afd3-7f8c39a2340c
	class ActionDataMoveCopy : public ActionData
	{
	protected:
		std::shared_ptr<blockT<BYTE>> FolderInThisStore = emptyT<BYTE>();
		std::shared_ptr<blockT<DWORD>> StoreEIDSize = emptyT<DWORD>();
		std::shared_ptr<EntryIdStruct> StoreEID;
		std::shared_ptr<blockBytes> StoreEIDBytes;
		std::shared_ptr<blockT<DWORD>> FolderEIDSize = emptyT<DWORD>();
		std::shared_ptr<ServerEID> FolderEID;
		std::shared_ptr<blockBytes> FolderEIDBytes;
		void parse() override
		{
			if (m_bExtended)
			{
				StoreEIDSize = blockT<DWORD>::parse(parser);
				StoreEIDBytes = blockBytes::parse(parser, *StoreEIDSize);
				FolderEIDSize = blockT<DWORD>::parse(parser);
				FolderEID = block::parse<ServerEID>(parser, *FolderEIDSize, true);
			}
			else
			{
				FolderInThisStore = blockT<BYTE>::parse(parser);
				StoreEIDSize = blockT<DWORD>::parse<WORD>(parser);
				if (*FolderInThisStore)
				{
					StoreEID = block::parse<EntryIdStruct>(parser, *StoreEIDSize, true);
				}
				else
				{
					StoreEIDBytes = blockBytes::parse(parser, *StoreEIDSize);
				}

				FolderEIDSize = blockT<DWORD>::parse<WORD>(parser);
				if (*FolderInThisStore)
				{
					FolderEID = block::parse<ServerEID>(parser, *FolderEIDSize, true);
				}
				else
				{
					FolderEIDBytes = blockBytes::parse(parser, *FolderEIDSize);
				}
			}
		}
		void parseBlocks()
		{
			setText(L"ActionDataMoveCopy:\r\n");
			if (m_bExtended)
			{
				addChild(StoreEIDSize);
				addChild(StoreEIDBytes);
				addChild(FolderEIDSize);
				addChild(FolderEID);
			}
			else
			{
				addChild(FolderInThisStore);
				addChild(StoreEIDSize);
				if (*FolderInThisStore)
				{
					addChild(StoreEID);
				}
				else
				{
					addChild(StoreEIDBytes);
				}

				addChild(FolderEIDSize);
				if (*FolderInThisStore)
				{
					addChild(FolderEID);
				}
				else
				{
					addChild(FolderEIDBytes);
				}
			}
		};
	};

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
		void parse() override { BounceCode = blockT<DWORD>::parse(parser); }

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
		setText(m_bExtended ? L"Extended Rule Action\r\n" : L"Rule Action\r\n");
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
} // namespace smartview