#include <StdAfx.h>
#include <Interpret/SmartView/EntryList.h>
#include <Interpret/String.h>

namespace smartview
{
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
					const auto cbRemainingBytes = min(m_Entry[i].EntryLength, m_Parser.RemainingBytes());
					m_Entry[i].EntryId.Init(cbRemainingBytes, m_Parser.GetCurrentAddress());
					m_Parser.Advance(cbRemainingBytes);
				}
			}
		}
	}

	_Check_return_ std::wstring EntryList::ToStringInternal()
	{
		auto szEntryList = strings::formatmessage(IDS_ENTRYLISTDATA, m_EntryCount, m_Pad);

		for (DWORD i = 0; i < m_Entry.size(); i++)
		{
			szEntryList +=
				strings::formatmessage(IDS_ENTRYLISTENTRYID, i, m_Entry[i].EntryLength, m_Entry[i].EntryLengthPad);

			szEntryList += m_Entry[i].EntryId.ToString();
		}

		return szEntryList;
	}
}