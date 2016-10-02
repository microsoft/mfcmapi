#include "stdafx.h"
#include "EntryList.h"
#include "String.h"

EntryList::EntryList(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_EntryCount = 0;
	m_Pad = 0;
	m_Entry = nullptr;
}

EntryList::~EntryList()
{
	if (m_Entry)
	{
		for (DWORD i = 0; i < m_EntryCount; i++)
		{
			delete m_Entry[i].EntryId;
		}
	}

	delete[] m_Entry;
}

void EntryList::Parse()
{
	m_Parser.GetDWORD(&m_EntryCount);
	m_Parser.GetDWORD(&m_Pad);

	if (m_EntryCount && m_EntryCount < _MaxEntriesLarge)
	{
		m_Entry = new EntryListEntryStruct[m_EntryCount];
		if (m_Entry)
		{
			memset(m_Entry, 0, sizeof(EntryListEntryStruct)* m_EntryCount);
			for (DWORD i = 0; i < m_EntryCount; i++)
			{
				m_Parser.GetDWORD(&m_Entry[i].EntryLength);
				m_Parser.GetDWORD(&m_Entry[i].EntryLengthPad);
			}

			for (DWORD i = 0; i < m_EntryCount; i++)
			{
				size_t cbRemainingBytes = min(m_Entry[i].EntryLength, m_Parser.RemainingBytes());
				m_Entry[i].EntryId = new EntryIdStruct(
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

	if (m_Entry)
	{
		for (DWORD i = 0; i < m_EntryCount; i++)
		{
			szEntryList += formatmessage(IDS_ENTRYLISTENTRYID,
				i,
				m_Entry[i].EntryLength,
				m_Entry[i].EntryLengthPad);

			if (m_Entry[i].EntryId)
			{
				szEntryList += m_Entry[i].EntryId->ToString();
			}
		}
	}

	return szEntryList;
}