#include <StdAfx.h>
#include <Interpret/SmartView/TombStone.h>
#include <Interpret/String.h>
#include <Interpret/SmartView/SmartView.h>

namespace smartview
{
	TombStone::TombStone() {}

	void TombStone::Parse()
	{
		m_Identifier = m_Parser.GetBlock<DWORD>();
		m_HeaderSize = m_Parser.GetBlock<DWORD>();
		m_Version = m_Parser.GetBlock<DWORD>();
		m_RecordsCount = m_Parser.GetBlock<DWORD>();
		m_RecordsSize = m_Parser.GetBlock<DWORD>();

		// Run through the parser once to count the number of flag structs
		const auto ulFlagOffset = m_Parser.GetCurrentOffset();
		for (;;)
		{
			// Must have at least 2 bytes left to have another flag
			if (m_Parser.RemainingBytes() < sizeof(DWORD) * 3 + sizeof(WORD)) break;
			(void) m_Parser.Get<DWORD>();
			(void) m_Parser.Get<DWORD>();
			m_Parser.Advance(m_Parser.Get<DWORD>());
			m_Parser.Advance(m_Parser.Get<WORD>());
			m_ActualRecordsCount++;
		}

		// Now we parse for real
		m_Parser.SetCurrentOffset(ulFlagOffset);

		if (m_ActualRecordsCount && m_ActualRecordsCount < _MaxEntriesSmall)
		{
			for (ULONG i = 0; i < m_ActualRecordsCount; i++)
			{
				TombstoneRecord tombstoneRecord;
				tombstoneRecord.StartTime = m_Parser.GetBlock<DWORD>();
				tombstoneRecord.EndTime = m_Parser.GetBlock<DWORD>();
				tombstoneRecord.GlobalObjectIdSize = m_Parser.GetBlock<DWORD>();
				tombstoneRecord.lpGlobalObjectId =
					m_Parser.GetBlockBYTES(tombstoneRecord.GlobalObjectIdSize.getData(), _MaxBytes);
				tombstoneRecord.UsernameSize = m_Parser.GetBlock<WORD>();
				tombstoneRecord.szUsername = m_Parser.GetBlockStringA(tombstoneRecord.UsernameSize.getData());
				m_lpRecords.push_back(tombstoneRecord);
			}
		}
	}

	_Check_return_ std::wstring TombStone::ToStringInternal()
	{
		auto szTombstoneString = strings::formatmessage(
			IDS_TOMBSTONEHEADER,
			m_Identifier.getData(),
			m_HeaderSize.getData(),
			m_Version.getData(),
			m_RecordsCount.getData(),
			m_ActualRecordsCount,
			m_RecordsSize.getData());

		for (ULONG i = 0; i < m_lpRecords.size(); i++)
		{
			auto szGoid = InterpretBinaryAsString(
				SBinary{static_cast<ULONG>(m_lpRecords[i].lpGlobalObjectId.getData().size()),
						m_lpRecords[i].lpGlobalObjectId.getData().data()},
				IDS_STGLOBALOBJECTID,
				nullptr);

			szTombstoneString += strings::formatmessage(
				IDS_TOMBSTONERECORD,
				i,
				m_lpRecords[i].StartTime.getData(),
				RTimeToString(m_lpRecords[i].StartTime.getData()).c_str(),
				m_lpRecords[i].EndTime.getData(),
				RTimeToString(m_lpRecords[i].EndTime.getData()).c_str(),
				m_lpRecords[i].GlobalObjectIdSize.getData(),
				strings::BinToHexString(m_lpRecords[i].lpGlobalObjectId.getData(), true).c_str(),
				szGoid.c_str(),
				m_lpRecords[i].UsernameSize.getData(),
				m_lpRecords[i].szUsername.getData().c_str());
		}

		return szTombstoneString;
	}
}