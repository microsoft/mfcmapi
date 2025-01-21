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
		StoreEntryId = blockBytes::parse(parser, *StoreEntryIdSize, _MaxEID);

		FolderEntryIdSize = blockT<DWORD>::parse(parser);
		FolderEntryId = blockBytes::parse(parser, *FolderEntryIdSize, _MaxEID);

		MessageEntryIdSize = blockT<DWORD>::parse(parser);
		MessageEntryId = blockBytes::parse(parser, *MessageEntryIdSize, _MaxEID);

		SearchFolderEntryIdSize = blockT<DWORD>::parse(parser);
		SearchFolderEntryId = blockBytes::parse(parser, *SearchFolderEntryIdSize, _MaxEID);

		MessageSearchKeySize = blockT<DWORD>::parse(parser);
		MessageSearchKey = blockBytes::parse(parser, *MessageSearchKeySize, _MaxEID);

		ANSITextSize = blockT<DWORD>::parse(parser);
		ANSIText = blockStringA::parse(parser, *ANSITextSize);
	}

	void ReportTag::addEID(
		const std::wstring& label,
		const std::shared_ptr<blockT<ULONG>>& _cb,
		const std::shared_ptr<blockBytes>& eid)
	{
		if (*_cb)
		{
			addLabeledChild(label, eid);
		}
	}

	void ReportTag::parseBlocks()
	{
		setText(L"Report Tag");
		addLabeledChild(L"Cookie", Cookie);

		auto szFlags = flags::InterpretFlags(flagReportTagVersion, *Version);
		addChild(Version, L"Version = 0x%1!08X! = %2!ws!", Version->getData(), szFlags.c_str());

		addEID(L"StoreEntryID", StoreEntryIdSize, StoreEntryId);
		addEID(L"FolderEntryID", FolderEntryIdSize, FolderEntryId);
		addEID(L"MessageEntryID", MessageEntryIdSize, MessageEntryId);
		addEID(L"SearchFolderEntryID", SearchFolderEntryIdSize, SearchFolderEntryId);
		addEID(L"MessageSearchKey", MessageSearchKeySize, MessageSearchKey);

		if (ANSITextSize)
		{
			addChild(ANSITextSize, L"cchAnsiText = 0x%1!08X!", ANSITextSize->getData());
			addChild(ANSIText, L"AnsiText = \"%1!hs!\"", ANSIText->c_str());
		}
	}
} // namespace smartview