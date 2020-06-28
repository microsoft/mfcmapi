#include <core/stdafx.h>
#include <core/smartview/EntryList.h>

namespace smartview
{
	void EntryListEntryStruct::parse()
	{
		EntryLength = blockT<DWORD>::parse(parser);
		EntryLengthPad = blockT<DWORD>::parse(parser);
	}

	void EntryListEntryStruct::parseEntryID()
	{
		EntryId = block::parse<EntryIdStruct>(parser, *EntryLength, true);

		addChild(EntryLength, L"EntryLength = 0x%1!08X!", EntryLength->getData());
		addChild(EntryLengthPad, L"EntryLengthPad = 0x%1!08X!", EntryLengthPad->getData());
		addChild(EntryId);
	}

	// We delay building our tree until we've finished parsing in parseEntryID
	void EntryListEntryStruct::parseBlocks() {}

	void EntryList::parse()
	{
		m_EntryCount = blockT<DWORD>::parse(parser);
		m_Pad = blockT<DWORD>::parse(parser);

		if (*m_EntryCount)
		{
			if (*m_EntryCount < _MaxEntriesLarge)
			{
				m_Entry.reserve(*m_EntryCount);
				for (DWORD i = 0; i < *m_EntryCount; i++)
				{
					m_Entry.emplace_back(block::parse<EntryListEntryStruct>(parser, false));
				}

				for (DWORD i = 0; i < *m_EntryCount; i++)
				{
					m_Entry[i]->parseEntryID();
				}
			}
		}
	}

	void EntryList::parseBlocks()
	{
		setText(L"Entry List");
		addChild(m_EntryCount, L"EntryCount = 0x%1!08X!", m_EntryCount->getData());
		addChild(m_Pad, L"Pad = 0x%1!08X!", m_Pad->getData());

		auto i = 0;
		for (const auto& entry : m_Entry)
		{
			addChild(entry, L"EntryId[%1!d!]", i);
			i++;
		}
	}
} // namespace smartview