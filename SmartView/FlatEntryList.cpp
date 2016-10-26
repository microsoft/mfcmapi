#include "stdafx.h"
#include "FlatEntryList.h"

FlatEntryList::FlatEntryList()
{
	m_cEntries = 0;
	m_cbEntries = 0;
}

void FlatEntryList::Parse()
{
	m_cEntries = m_Parser.Get<DWORD>();

	// We read and report this, but ultimately, it's not used.
	m_cbEntries = m_Parser.Get<DWORD>();

	if (m_cEntries && m_cEntries < _MaxEntriesLarge)
	{
		for (DWORD iFlatEntryList = 0; iFlatEntryList < m_cEntries; iFlatEntryList++)
		{
			FlatEntryID flatEntryID;
			// Size here will be the length of the serialized entry ID
			// We'll have to round it up to a multiple of 4 to read off padding
			flatEntryID.dwSize = m_Parser.Get<DWORD>();
			size_t ulSize = min(flatEntryID.dwSize, m_Parser.RemainingBytes());

			flatEntryID.lpEntryID.Init(
				static_cast<ULONG>(ulSize),
				m_Parser.GetCurrentAddress());

			m_Parser.Advance(ulSize);

			auto dwPAD = 3 - (flatEntryID.dwSize + 3) % 4;
			if (dwPAD > 0)
			{
				flatEntryID.JunkData = m_Parser.GetBYTES(dwPAD);
			}

			m_pEntryIDs.push_back(flatEntryID);
		}
	}
}

_Check_return_ wstring FlatEntryList::ToStringInternal()
{
	wstring szFlatEntryList;

	szFlatEntryList = formatmessage(
		IDS_FELHEADER,
		m_cEntries,
		m_cbEntries);

	for (DWORD iFlatEntryList = 0; iFlatEntryList < m_pEntryIDs.size(); iFlatEntryList++)
	{
		szFlatEntryList += formatmessage(
			IDS_FELENTRYHEADER,
			iFlatEntryList,
			m_pEntryIDs[iFlatEntryList].dwSize);
		szFlatEntryList += m_pEntryIDs[iFlatEntryList].lpEntryID.ToString();

		if (m_pEntryIDs[iFlatEntryList].JunkData.size())
		{
			szFlatEntryList += formatmessage(
				IDS_FELENTRYPADDING,
				iFlatEntryList);
			szFlatEntryList += JunkDataToString(m_pEntryIDs[iFlatEntryList].JunkData);
		}

		szFlatEntryList += L"\r\n"; // STRING_OK
	}

	return szFlatEntryList;
}