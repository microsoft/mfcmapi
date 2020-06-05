#include <core/stdafx.h>
#include <core/smartview/EntryList.h>
#include <core/smartview/block/scratchBlock.h>

namespace smartview
{
	EntryListEntryStruct::EntryListEntryStruct(const std::shared_ptr<binaryParser>& parser)
	{
		EntryLength = blockT<DWORD>::parse(parser);
		EntryLengthPad = blockT<DWORD>::parse(parser);
	}

	void EntryList::parse()
	{
		m_EntryCount = blockT<DWORD>::parse(m_Parser);
		m_Pad = blockT<DWORD>::parse(m_Parser);

		if (*m_EntryCount)
		{
			if (*m_EntryCount < _MaxEntriesLarge)
			{
				m_Entry.reserve(*m_EntryCount);
				for (DWORD i = 0; i < *m_EntryCount; i++)
				{
					m_Entry.emplace_back(std::make_shared<EntryListEntryStruct>(m_Parser));
				}

				for (DWORD i = 0; i < *m_EntryCount; i++)
				{
					m_Entry[i]->EntryId =
						smartViewParser::parse<EntryIdStruct>(m_Parser, *m_Entry[i]->EntryLength, true);
				}
			}
		}
	}

	void EntryList::parseBlocks()
	{
		addChild(m_EntryCount, L"EntryCount = 0x%1!08X!\r\n", m_EntryCount->getData());
		addChild(m_Pad, L"Pad = 0x%1!08X!", m_Pad->getData());

		auto i = 0;
		for (const auto& entry : m_Entry)
		{
			terminateBlock();
			addHeader(L"EntryId[%1!d!]:\r\n", i);
			addChild(entry->EntryLength, L"EntryLength = 0x%1!08X!\r\n", entry->EntryLength->getData());
			addChild(entry->EntryLengthPad, L"EntryLengthPad = 0x%1!08X!\r\n", entry->EntryLengthPad->getData());
			addChild(labeledBlock(L"Entry Id = ", entry->EntryId));

			i++;
		}
	}
} // namespace smartview