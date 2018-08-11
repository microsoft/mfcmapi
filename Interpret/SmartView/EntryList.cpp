#include <StdAfx.h>
#include <Interpret/SmartView/EntryList.h>
#include <Interpret/String.h>

namespace smartview
{
	EntryList::EntryList() {}

	void EntryList::Parse()
	{
		m_EntryCount = m_Parser.GetBlock<DWORD>();
		m_Pad = m_Parser.GetBlock<DWORD>();

		if (m_EntryCount.getData() && m_EntryCount.getData() < _MaxEntriesLarge)
		{
			{
				for (DWORD i = 0; i < m_EntryCount.getData(); i++)
				{
					EntryListEntryStruct entryListEntryStruct;
					entryListEntryStruct.EntryLength = m_Parser.GetBlock<DWORD>();
					entryListEntryStruct.EntryLengthPad = m_Parser.GetBlock<DWORD>();
					m_Entry.push_back(entryListEntryStruct);
				}

				for (DWORD i = 0; i < m_EntryCount.getData(); i++)
				{
					const auto cbRemainingBytes = min(m_Entry[i].EntryLength.getData(), m_Parser.RemainingBytes());
					m_Entry[i].EntryId.Init(cbRemainingBytes, m_Parser.GetCurrentAddress());
					m_Entry[i].EntryId.EnsureParsed();
					m_Parser.Advance(cbRemainingBytes);
				}
			}
		}
	}

	void EntryList::ParseBlocks()
	{
		addBlock(m_EntryCount, L"EntryCount = 0x%1!08X!\r\n", m_EntryCount.getData());
		addBlock(m_Pad, L"Pad = 0x%1!08X!", m_Pad.getData());

		for (DWORD i = 0; i < m_Entry.size(); i++)
		{
			addLine();
			addHeader(L"EntryId[%1!d!]:\r\n", i);
			addBlock(m_Entry[i].EntryLength, L"EntryLength = 0x%1!08X!\r\n", m_Entry[i].EntryLength.getData());
			addBlock(m_Entry[i].EntryLengthPad, L"EntryLengthPad = 0x%1!08X!\r\n", m_Entry[i].EntryLengthPad.getData());
			addHeader(L"Entry Id = ");
			addBlock(m_Entry[i].EntryId.getBlock());
		}
	}
}