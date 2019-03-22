#include <core/stdafx.h>
#include <core/smartview/EntryList.h>

namespace smartview
{
	EntryListEntryStruct::EntryListEntryStruct(std::shared_ptr<binaryParser> parser)
	{
		EntryLength = blockT<DWORD>::parse(parser);
		EntryLengthPad = blockT<DWORD>::parse(parser);
	}

	void EntryList::Parse()
	{
		m_EntryCount = blockT<DWORD>::parse(m_Parser);
		m_Pad = blockT<DWORD>::parse(m_Parser);

		if (*m_EntryCount)
		{
			if (*m_EntryCount < _MaxEntriesLarge)
			{
				m_Entry.reserve(*m_EntryCount);
				for (auto i = 0; i < *m_EntryCount; i++)
				{
					m_Entry.emplace_back(std::make_shared<EntryListEntryStruct>(m_Parser));
				}

				for (DWORD i = 0; i < *m_EntryCount; i++)
				{
					m_Entry[i]->EntryId.parse(m_Parser, *m_Entry[i]->EntryLength, true);
				}
			}
		}
	}

	void EntryList::ParseBlocks()
	{
		setRoot(m_EntryCount, L"EntryCount = 0x%1!08X!\r\n", m_EntryCount->getData());
		addChild(m_Pad, L"Pad = 0x%1!08X!", m_Pad->getData());

		auto i = 0;
		for (const auto& entry : m_Entry)
		{
			terminateBlock();
			addHeader(L"EntryId[%1!d!]:\r\n", i);
			addChild(entry->EntryLength, L"EntryLength = 0x%1!08X!\r\n", entry->EntryLength->getData());
			addChild(entry->EntryLengthPad, L"EntryLengthPad = 0x%1!08X!\r\n", entry->EntryLengthPad->getData());
			addHeader(L"Entry Id = ");
			addChild(entry->EntryId.getBlock());

			i++;
		}
	}
} // namespace smartview