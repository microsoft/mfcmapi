#include <core/stdafx.h>
#include <core/smartview/ReportTag.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void ReportTag::parse()
	{
		Cookie = blockBytes::parse(parser, 9);

		Version = blockT<DWORD>::parse(parser);

		StoreEntryIdSize = blockT<DWORD>::parse(parser);
		if (*StoreEntryIdSize > 0)
		{
			StoreEntryId = block::parse<EntryIdStruct>(parser, *StoreEntryIdSize, true);
		}

		FolderEntryIdSize = blockT<DWORD>::parse(parser);
		if (*FolderEntryIdSize > 0)
		{
			FolderEntryId = block::parse<EntryIdStruct>(parser, *FolderEntryIdSize, true);
		}

		MessageEntryIdSize = blockT<DWORD>::parse(parser);
		if (*MessageEntryIdSize > 0)
		{
			MessageEntryId = block::parse<EntryIdStruct>(parser, *MessageEntryIdSize, true);
		}

		SearchFolderEntryIdSize = blockT<DWORD>::parse(parser);
		if (*SearchFolderEntryIdSize > 0)
		{
			SearchFolderEntryId = block::parse<EntryIdStruct>(parser, *SearchFolderEntryIdSize, true);
		}

		MessageSearchKeySize = blockT<DWORD>::parse(parser);
		if (*MessageSearchKeySize > 0)
		{
			MessageSearchKey = blockBytes::parse(parser, *MessageSearchKeySize, _MaxEID);
		}

		ANSITextSize = blockT<DWORD>::parse(parser);
		ANSIText = blockStringA::parse(parser, *ANSITextSize);
	}

	void ReportTag::parseBlocks()
	{
		setText(L"Report Tag");
		addLabeledChild(L"Cookie", Cookie);

		auto szFlags = flags::InterpretFlags(flagReportTagVersion, *Version);
		addChild(Version, L"Version = 0x%1!08X! = %2!ws!", Version->getData(), szFlags.c_str());

		addChild(StoreEntryIdSize, L"StoreEntryIdSize = 0x%1!08X!", StoreEntryIdSize->getData());
		addChild(StoreEntryId, L"StoreEntryID");
		addChild(FolderEntryIdSize, L"FolderEntryIdSize = 0x%1!08X!", FolderEntryIdSize->getData());
		addChild(FolderEntryId, L"FolderEntryID");
		addChild(MessageEntryIdSize, L"MessageEntryIdSize = 0x%1!08X!", MessageEntryIdSize->getData());
		addChild(MessageEntryId, L"MessageEntryID");
		addChild(SearchFolderEntryIdSize, L"SearchFolderEntryIdSize = 0x%1!08X!", SearchFolderEntryIdSize->getData());
		addChild(SearchFolderEntryId, L"SearchFolderEntryID");
		addChild(MessageSearchKeySize, L"MessageSearchKeySize = 0x%1!08X!", MessageSearchKeySize->getData());
		addLabeledChild(L"MessageSearchKey", MessageSearchKey);

		if (ANSITextSize)
		{
			addChild(ANSITextSize, L"cchAnsiText = 0x%1!08X!", ANSITextSize->getData());
			addChild(ANSIText, L"AnsiText = \"%1!hs!\"", ANSIText->c_str());
		}
	}
} // namespace smartview