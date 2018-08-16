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

		if (m_cEntries && m_cEntries < _MaxEntriesLarge)
		{
			for (DWORD iFlatEntryList = 0; iFlatEntryList < m_cEntries; iFlatEntryList++)
			{
				FlatEntryID flatEntryID;
				// Size here will be the length of the serialized entry ID
				// We'll have to round it up to a multiple of 4 to read off padding
				flatEntryID.dwSize = m_Parser.GetBlock<DWORD>();
				const auto ulSize = min(flatEntryID.dwSize, m_Parser.RemainingBytes());

				flatEntryID.lpEntryID.Init(ulSize, m_Parser.GetCurrentAddress());
				flatEntryID.lpEntryID.EnsureParsed();

				m_Parser.Advance(ulSize);

				const auto dwPAD = 3 - (flatEntryID.dwSize + 3) % 4;
				if (dwPAD > 0)
				{
					flatEntryID.JunkData = m_Parser.GetBlockBYTES(dwPAD);
				}

				m_pEntryIDs.push_back(flatEntryID);
			}
		}
	}

	void FlatEntryList::ParseBlocks()
	{
		addHeader(L"Flat Entry List\r\n");
		addBlock(m_cEntries, L"cEntries = %1!d!\r\n", m_cEntries.getData());
		addBlock(m_cbEntries, L"cbEntries = 0x%1!08X!", m_cbEntries.getData());

		for (DWORD iFlatEntryList = 0; iFlatEntryList < m_pEntryIDs.size(); iFlatEntryList++)
		{
			addLine();
			addLine();
			addHeader(L"Entry[%1!d!] ", iFlatEntryList);
			addBlock(
				m_pEntryIDs[iFlatEntryList].dwSize,
				L"Size = 0x%1!08X!",
				m_pEntryIDs[iFlatEntryList].dwSize.getData());

			if (m_pEntryIDs[iFlatEntryList].lpEntryID.hasData())
			{
				addLine();
				addBlock(m_pEntryIDs[iFlatEntryList].lpEntryID.getBlock());
			}

			if (m_pEntryIDs[iFlatEntryList].JunkData.getData().size())
			{
				addLine();
				addHeader(L"Entry[%1!d!] Padding:", iFlatEntryList);
				addHeader(JunkDataToString(m_pEntryIDs[iFlatEntryList].JunkData.getData()));
			}
		}
	}
}