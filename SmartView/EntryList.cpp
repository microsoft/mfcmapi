#include "stdafx.h"
#include "EntryList.h"
#include "String.h"

EntryList::EntryList()
{
	m_EntryCount = 0;
	m_Pad = 0;
}

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
				size_t cbRemainingBytes = min(m_Entry[i].EntryLength, m_Parser.RemainingBytes());
				m_Entry[i].EntryId.Init(
					static_cast<ULONG>(cbRemainingBytes),
					m_Parser.GetCurrentAddress());
				m_Parser.Advance(cbRemainingBytes);
			}
		}
	}
}

_Check_return_ wstring EntryList::ToStringInternal()
{
	auto szEntryList = formatmessage(IDS_ENTRYLISTDATA,
		m_EntryCount,
		m_Pad);

	for (DWORD i = 0; i < m_EntryCount; i++)
	{
		szEntryList += formatmessage(IDS_ENTRYLISTENTRYID,
			i,
			m_Entry[i].EntryLength,
			m_Entry[i].EntryLengthPad);

		szEntryList += m_Entry[i].EntryId.ToString();
	}

	return szEntryList;
}