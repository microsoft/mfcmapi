#include <StdAfx.h>
#include <Interpret/SmartView/EntryList.h>
#include <Interpret/String.h>

namespace smartview
{
	void EntryList::Parse()
	{
		m_EntryCount = m_Parser.Get<DWORD>();
		m_Pad = m_Parser.Get<DWORD>();

		if (m_EntryCount && m_EntryCount < _MaxEntriesLarge)
		{
			{
				for (DWORD i = 0; i < m_EntryCount; i++)
				{
					EntryListEntryStruct entryListEntryStruct;
					entryListEntryStruct.EntryLength = m_Parser.Get<DWORD>();
					entryListEntryStruct.EntryLengthPad = m_Parser.Get<DWORD>();
					m_Entry.push_back(entryListEntryStruct);
				}

				for (DWORD i = 0; i < m_EntryCount; i++)
				{
					const auto cbRemainingBytes = min(m_Entry[i].EntryLength, m_Parser.RemainingBytes());
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