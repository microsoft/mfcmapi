#include <core/stdafx.h>
#include <core/smartview/RuleAction.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/interpret/guid.h>

namespace smartview
{
	// 2.2.5.1.2.1.1 ServerEid Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/07e8c314-0ab2-440e-9138-b96f93682bf1
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

	// 2.2.5.1.2.2 OP_REPLY and OP_OOF_REPLY ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/d5c14d2f-3557-4b78-acfe-d949ef325792
	class ActionDataReply : public ActionData
	{
	protected:
		std::shared_ptr<blockT<LARGE_INTEGER>> ReplyTemplateFID = emptyT<LARGE_INTEGER>();
		std::shared_ptr<blockT<LARGE_INTEGER>> ReplyTemplateMID = emptyT<LARGE_INTEGER>();
		std::shared_ptr<blockT<DWORD>> MessageEIDSize = emptyT<DWORD>();
		std::shared_ptr<EntryIdStruct> ReplyTemplateMessageEID;
		std::shared_ptr<blockT<GUID>> ReplyTemplateGUID = emptyT<GUID>();
		void parse() override
		{
			if (m_bExtended)
			{
				MessageEIDSize = blockT<DWORD>::parse(parser);
				ReplyTemplateMessageEID = block::parse<EntryIdStruct>(parser, *MessageEIDSize, true);
			}
			else
			{
				ReplyTemplateFID = blockT<LARGE_INTEGER>::parse(parser);
				ReplyTemplateMID = blockT<LARGE_INTEGER>::parse(parser);
			}

			ReplyTemplateGUID = blockT<GUID>::parse(parser);
		}
		void parseBlocks()
		{
			setText(L"ActionDataReply:\r\n");
			if (m_bExtended)
			{
				addChild(MessageEIDSize);
				addChild(ReplyTemplateMessageEID);
				addChild(ReplyTemplateGUID);
			}
			else
			{
				addChild(ReplyTemplateFID);
				addChild(ReplyTemplateMID);
				addChild(ReplyTemplateGUID);
			}
		};
	};

	// 2.2.5.1.2.3 OP_DEFER_ACTION ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/1d66ff31-fdc1-4b36-a040-75207e31dd18
	class ActionDataDefer : public ActionData
	{
	protected:
		void parse() override {}
		void parseBlocks() { setText(L"ActionDataDefer:\r\n"); };
	};

	// 2.2.5.1.2.4.1 RecipientBlockData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/c6ad133d-7906-43aa-8420-3b40ac6be494
	class RecipientBlockData : public block
	{
	private:
		std::shared_ptr<blockT<BYTE>> Reserved = emptyT<BYTE>();
		std::shared_ptr<blockT<DWORD>> NoOfProperties = emptyT<DWORD>();
		std::vector<std::shared_ptr<SPropValueStruct>> PropertyValues;

		void parse() override
		{
			Reserved = blockT<BYTE>::parse(parser);
			NoOfProperties = blockT<DWORD>::parse(parser);
			if (*NoOfProperties && *NoOfProperties < _MaxEntriesLarge)
			{
				PropertyValues.reserve(*NoOfProperties);
				for (DWORD i = 0; i < *NoOfProperties; i++)
				{
					PropertyValues.push_back(block::parse<SPropValueStruct>(parser, 0, true));
				}
			}
		};
		void parseBlocks() override
		{
			setText(L"RecipientBlockData:\r\n");
			addChild(Reserved);
			addChild(NoOfProperties);
			for (const auto& propertyValue : PropertyValues)
			{
				addChild(propertyValue);
				terminateBlock();
			}
		};
	};

	// 2.2.5.1.2.4 OP_FORWARD and OP_DELEGATE ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/1b8a2b3f-9fff-4e07-9820-c3d2287a8e5c
	class ActionDataForwardDelegate : public ActionData
	{
	protected:
		std::shared_ptr<blockT<DWORD>> RecipientCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<RecipientBlockData>> RecipientBlocks;
		void parse() override
		{
			RecipientCount = blockT<DWORD>::parse(parser);
			if (*RecipientCount && *RecipientCount < _MaxEntriesLarge)
			{
				RecipientBlocks.reserve(*RecipientCount);
				for (DWORD i = 0; i < *RecipientCount; i++)
				{
					RecipientBlocks.push_back(block::parse<RecipientBlockData>(parser, 0, true));
				}
			}
		}
		void parseBlocks()
		{
			setText(L"ActionDataForwardDelegate:\r\n");
			addChild(RecipientCount);
			for (const auto& recipient : RecipientBlocks)
			{
				addChild(recipient);
				terminateBlock();
			}
		};
	};

	// 2.2.5.1.2.5 OP_BOUNCE ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/c6ceb0c2-96a2-4337-a4fb-f2e23cfe4284
	class ActionDataBounce : public ActionData
	{
	private:
		std::shared_ptr<blockT<DWORD>> BounceCode = emptyT<DWORD>();
		void parse() override { BounceCode = blockT<DWORD>::parse(parser); }

		void parseBlocks()
		{
			setText(L"ActionDataBounce:\r\n");
			addChild(BounceCode);
		}
	};

	// 2.2.5.1.2.6 OP_TAG ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/2fe6b110-5e1e-4a97-aeeb-9103cf76a0e0
	class ActionDataTag : public ActionData
	{
	private:
		std::shared_ptr<SPropValueStruct> TaggedPropertyValue;
		void parse() override { TaggedPropertyValue = block::parse<SPropValueStruct>(parser, false); }

		void parseBlocks()
		{
			setText(L"ActionDataTag:\r\n");
			addChild(TaggedPropertyValue);
		}
	};

	// 2.2.5.1.2.7 OP_DELETE or OP_MARK_AS_READ ActionData Structure
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/c180c7cb-474a-46f2-95ee-568ec61f1043
	class ActionDataDeleteMarkRead : public ActionData
	{
	private:
		void parse() override {}
		void parseBlocks() { setText(L"ActionDataDeleteMarkRead:\r\n"); }
	};

	std::shared_ptr<ActionData> getActionDataParser(BYTE at)
	{
		switch (at)
		{
		case OP_MOVE:
			return std::make_shared<ActionDataMoveCopy>();
			break;
		case OP_COPY:
			return std::make_shared<ActionDataMoveCopy>();
			break;
		case OP_REPLY:
			return std::make_shared<ActionDataReply>();
			break;
		case OP_OOF_REPLY:
			return std::make_shared<ActionDataReply>();
			break;
		case OP_DEFER_ACTION:
			return std::make_shared<ActionDataDefer>();
			break;
		case OP_BOUNCE:
			return std::make_shared<ActionDataBounce>();
			break;
		case OP_FORWARD:
			return std::make_shared<ActionDataForwardDelegate>();
			break;
		case OP_DELEGATE:
			return std::make_shared<ActionDataForwardDelegate>();
			break;
		case OP_TAG:
			return std::make_shared<ActionDataTag>();
			break;
		case OP_DELETE:
			return std::make_shared<ActionDataDeleteMarkRead>();
			break;
		case OP_MARK_AS_READ:
			return std::make_shared<ActionDataDeleteMarkRead>();
			break;
		}

		return {};
	}

	void ActionBlock ::parse()
	{
		if (m_bExtended)
		{
			ActionLength = blockT<DWORD>::parse(parser);
		}
		else
		{
			ActionLength = blockT<DWORD>::parse<WORD>(parser);
		}

		ActionType = blockT<BYTE>::parse(parser);
		ActionFlavor = blockT<DWORD>::parse(parser);
		ActionData = getActionDataParser(*ActionType);
		if (ActionData)
		{
			ActionData->parse(parser, m_bExtended);
		}

	}
	void ActionBlock ::parseBlocks() {}

} // namespace smartview