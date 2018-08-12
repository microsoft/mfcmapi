#include <StdAfx.h>
#include <Interpret/SmartView/FlatEntryList.h>

namespace smartview
{
	FlatEntryList::FlatEntryList() {}

	void FlatEntryList::Parse()
	{
		m_cEntries = m_Parser.GetBlock<DWORD>();

		// We read and report this, but ultimately, it's not used.
		m_cbEntries = m_Parser.GetBlock<DWORD>();

		if (m_cEntries.getData() && m_cEntries.getData() < _MaxEntriesLarge)
		{
			for (DWORD iFlatEntryList = 0; iFlatEntryList < m_cEntries.getData(); iFlatEntryList++)
			{
				FlatEntryID flatEntryID;
				// Size here will be the length of the serialized entry ID
				// We'll have to round it up to a multiple of 4 to read off padding
				flatEntryID.dwSize = m_Parser.GetBlock<DWORD>();
				const auto ulSize = min(flatEntryID.dwSize.getData(), m_Parser.RemainingBytes());

				flatEntryID.lpEntryID.Init(ulSize, m_Parser.GetCurrentAddress());

				m_Parser.Advance(ulSize);

				const auto dwPAD = 3 - (flatEntryID.dwSize.getData() + 3) % 4;
				if (dwPAD > 0)
				{
					flatEntryID.JunkData = m_Parser.GetBlockBYTES(dwPAD);
				}

				m_pEntryIDs.push_back(flatEntryID);
			}
		}
	}

	_Check_return_ std::wstring FlatEntryList::ToStringInternal()
	{
		std::vector<std::wstring> flatEntryList;
		flatEntryList.push_back(strings::formatmessage(IDS_FELHEADER, m_cEntries.getData(), m_cbEntries.getData()));

		for (DWORD iFlatEntryList = 0; iFlatEntryList < m_pEntryIDs.size(); iFlatEntryList++)
		{
			flatEntryList.push_back(strings::formatmessage(
				IDS_FELENTRYHEADER, iFlatEntryList, m_pEntryIDs[iFlatEntryList].dwSize.getData()));
			auto entryID = m_pEntryIDs[iFlatEntryList].lpEntryID.ToString();

			if (entryID.length())
			{
				flatEntryList.push_back(entryID);
			}

			if (m_pEntryIDs[iFlatEntryList].JunkData.getData().size())
			{
				flatEntryList.push_back(
					strings::formatmessage(IDS_FELENTRYPADDING, iFlatEntryList) +
					JunkDataToString(m_pEntryIDs[iFlatEntryList].JunkData.getData()));
			}
		}

		return strings::join(flatEntryList, L"\r\n"); //STRING_OK
	}
}