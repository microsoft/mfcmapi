#include "stdafx.h"
#include "..\stdafx.h"
#include "EntryList.h"
#include "..\String.h"

EntryList::EntryList(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_EntryCount = 0;
	m_Pad = 0;
	m_Entry = NULL;
}

EntryList::~EntryList()
{
	if (m_Entry)
	{
		DWORD i = 0;
		for (i = 0; i < m_EntryCount; i++)
		{
			DeleteEntryIdStruct(m_Entry[i].EntryId);
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
			DWORD i = 0;
			for (i = 0; i < m_EntryCount; i++)
			{
				m_Parser.GetDWORD(&m_Entry[i].EntryLength);
				m_Parser.GetDWORD(&m_Entry[i].EntryLengthPad);
			}

			for (i = 0; i < m_EntryCount; i++)
			{
				size_t cbRemainingBytes = m_Parser.RemainingBytes();
				cbRemainingBytes = min(m_Entry[i].EntryLength, cbRemainingBytes);
				m_Entry[i].EntryId = BinToEntryIdStruct(
					(ULONG)cbRemainingBytes,
					m_Parser.GetCurrentAddress());
				m_Parser.Advance(cbRemainingBytes);
			}
		}
	}
}

_Check_return_ wstring EntryList::ToStringInternal()
{
	wstring szEntryList;
	wstring szTmp;

	szEntryList = formatmessage(IDS_ENTRYLISTDATA,
		m_EntryCount,
		m_Pad);

	if (m_Entry)
	{
		DWORD i = m_EntryCount;
		for (i = 0; i < m_EntryCount; i++)
		{
			szTmp = formatmessage(IDS_ENTRYLISTENTRYID,
				i,
				m_Entry[i].EntryLength,
				m_Entry[i].EntryLengthPad);
			szEntryList += szTmp;
			LPWSTR szEntryId = EntryIdStructToString(m_Entry[i].EntryId);
			szEntryList += szEntryId;
			delete[] szEntryId;
		}
	}

	return szEntryList;
}