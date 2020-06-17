#include <core/stdafx.h>
#include <core/smartview/RuleAction.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

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
			addChild(Ours, L"Ours: %1!ws!\r\n", *Ours ? L"true" : L"false");
			addChild(
				FolderId,
				L"FolderId: 0x%1!08X!:0x%2!08X!\r\n",
				FolderId->getData().HighPart,
				FolderId->getData().LowPart);
			addChild(
				MessageId,
				L"MessageId: 0x%1!08X!:0x%2!08X!\r\n",
				MessageId->getData().HighPart,
				MessageId->getData().LowPart);
			addChild(Instance, L"Instance: 0x%1!04X!\r\n", Instance->getData());
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
		std::shared_ptr<EntryIdStruct> FolderEID;
		std::shared_ptr<ServerEID> FolderEIDserverEID;
		std::shared_ptr<blockBytes> FolderEIDBytes;
		void parse() override
		{
			if (m_bExtended)
			{
				FolderInThisStore = blockT<BYTE>::parse(parser);
				StoreEIDSize = blockT<DWORD>::parse(parser);
				StoreEIDBytes = blockBytes::parse(parser, *StoreEIDSize);
				FolderEIDSize = blockT<DWORD>::parse(parser);
				FolderEID = block::parse<EntryIdStruct>(parser, *FolderEIDSize, true);
			}
			else
			{
				FolderInThisStore = blockT<BYTE>::parse(parser);
				StoreEIDSize = blockT<DWORD>::parse<WORD>(parser);
				if (*FolderInThisStore)
				{
					StoreEIDBytes = blockBytes::parse(parser, *StoreEIDSize);
				}
				else
				{
					StoreEID = block::parse<EntryIdStruct>(parser, *StoreEIDSize, true);
				}

				FolderEIDSize = blockT<DWORD>::parse<WORD>(parser);
				if (*FolderInThisStore)
				{
					FolderEIDserverEID = block::parse<ServerEID>(parser, *FolderEIDSize, true);
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
				addChild(FolderInThisStore, L"FolderInThisStore: %1!ws!\r\n", *FolderInThisStore ? L"true" : L"false");
				addChild(StoreEIDSize, L"StoreEIDSize: 0x%1!08X!\r\n", StoreEIDSize->getData());
				addLabeledChild(L"StoreEIDBytes = ", StoreEIDBytes);
				addChild(FolderEIDSize, L"FolderEIDSize: 0x%1!08X!\r\n", FolderEIDSize->getData());
				addLabeledChild(L"FolderEID = ", FolderEID);
			}
			else
			{
				addChild(FolderInThisStore, L"FolderInThisStore: %1!ws!\r\n", *FolderInThisStore ? L"true" : L"false");
				addChild(StoreEIDSize, L"StoreEIDSize: 0x%1!04X!\r\n", StoreEIDSize->getData());
				if (*FolderInThisStore)
				{
					addChild(StoreEID, L"StoreEID\r\n");
				}
				else
				{
					addLabeledChild(L"StoreEIDBytes = ", StoreEIDBytes);
				}

				addChild(FolderEIDSize, L"FolderEIDSize: 0x%1!04X!\r\n", FolderEIDSize->getData());
				if (*FolderInThisStore)
				{
					addLabeledChild(L"FolderEIDserverEID = ", FolderEIDserverEID);
				}
				else
				{
					addLabeledChild(L"FolderEIDBytes = ", FolderEIDBytes);
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
				addChild(MessageEIDSize, L"MessageEIDSize: 0x%1!08X!\r\n", MessageEIDSize->getData());
				addChild(ReplyTemplateMessageEID, L"ReplyTemplateMessageEID");
			}
			else
			{
				addChild(
					ReplyTemplateFID,
					L"ReplyTemplateFID: 0x%1!08X!:0x%2!08X!\r\n",
					ReplyTemplateFID->getData().HighPart,
					ReplyTemplateFID->getData().LowPart);
				addChild(
					ReplyTemplateMID,
					L"ReplyTemplateMID: 0x%1!08X!:0x%2!08X!\r\n",
					ReplyTemplateMID->getData().HighPart,
					ReplyTemplateMID->getData().LowPart);
			}

			addChild(
				ReplyTemplateGUID,
				L"ReplyTemplateGUID = %1!ws!",
				guid::GUIDToStringAndName(*ReplyTemplateGUID).c_str());
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
			addChild(Reserved, L"Reserved: 0x%1!01X!\r\n", Reserved->getData());
			addChild(NoOfProperties, L"NoOfProperties: 0x%1!08X!\r\n", NoOfProperties->getData());
			auto i = 0;
			for (const auto& propertyValue : PropertyValues)
			{
				addChild(propertyValue, L"PropertyValues[%1!d!]\r\n", i++);
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
			addChild(RecipientCount, L"RecipientCount: 0x%1!08X!\r\n", RecipientCount->getData());
			auto i = 0;
			for (const auto& recipient : RecipientBlocks)
			{
				addChild(recipient, L"RecipientBlocks[%1!d!]\r\n", i++);
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
			addChild(BounceCode, L"BounceCode: 0x%1!08X!\r\n", BounceCode->getData());
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
			addChild(TaggedPropertyValue, L"TaggedPropertyValue");
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

	std::shared_ptr<ActionData> getActionDataParser(DWORD at)
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

		ActionType = blockT<DWORD>::parse(parser); // TODO: not m_bExtended may be 1 byte
		ActionFlavor = blockT<DWORD>::parse(parser);
		ActionData = getActionDataParser(*ActionType);
		if (ActionData)
		{
			ActionData->parse(parser, m_bExtended);
		}
	}

	void ActionBlock ::parseBlocks()
	{
		setText(L"ActionBlock:\r\n");
		addChild(ActionLength, L"ActionLength: 0x%1!08X!\r\n", ActionLength->getData());
		addChild(
			ActionType,
			L"ActionType: %1!ws!\r\n",
			flags::InterpretFlags(flagActionType, ActionType->getData()).c_str());
		addChild(ActionFlavor, L"ActionFlavor: 0x%1!08X!\r\n", ActionFlavor->getData());
		addChild(ActionData);
	}
} // namespace smartview