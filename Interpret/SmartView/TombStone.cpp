#include <StdAfx.h>
#include <Interpret/SmartView/TombStone.h>
#include <Interpret/String.h>
#include <Interpret/SmartView/SmartView.h>

namespace smartview
{
	TombStone::TombStone()
	{
		m_Identifier = 0;
		m_HeaderSize = 0;
		m_Version = 0;
		m_RecordsCount = 0;
		m_ActualRecordsCount = 0;
		m_RecordsSize = 0;
	}

	void TombStone::Parse()
	{
		m_Identifier = m_Parser.Get<DWORD>();
		m_HeaderSize = m_Parser.Get<DWORD>();
		m_Version = m_Parser.Get<DWORD>();
		m_RecordsCount = m_Parser.Get<DWORD>();
		m_RecordsSize = m_Parser.Get<DWORD>();

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
				tombstoneRecord.StartTime = m_Parser.Get<DWORD>();
				tombstoneRecord.EndTime = m_Parser.Get<DWORD>();
				tombstoneRecord.GlobalObjectIdSize = m_Parser.Get<DWORD>();
				tombstoneRecord.lpGlobalObjectId = m_Parser.GetBYTES(tombstoneRecord.GlobalObjectIdSize, _MaxBytes);
				tombstoneRecord.UsernameSize = m_Parser.Get<WORD>();
				tombstoneRecord.szUsername = m_Parser.GetStringA(tombstoneRecord.UsernameSize);
				m_lpRecords.push_back(tombstoneRecord);
			}
		}
	}

	_Check_return_ std::wstring TombStone::ToStringInternal()
	{
		auto szTombstoneString = strings::formatmessage(
			IDS_TOMBSTONEHEADER,
			m_Identifier,
			m_HeaderSize,
			m_Version,
			m_RecordsCount,
			m_ActualRecordsCount,
			m_RecordsSize);

		for (ULONG i = 0; i < m_lpRecords.size(); i++)
		{
			const SBinary sBin = {static_cast<ULONG>(m_lpRecords[i].lpGlobalObjectId.size()),
								  m_lpRecords[i].lpGlobalObjectId.data()};

			auto szGoid = InterpretBinaryAsString(sBin, IDS_STGLOBALOBJECTID, nullptr);

			szTombstoneString += strings::formatmessage(
				IDS_TOMBSTONERECORD,
				i,
				m_lpRecords[i].StartTime,
				RTimeToString(m_lpRecords[i].StartTime).c_str(),
				m_lpRecords[i].EndTime,
				RTimeToString(m_lpRecords[i].EndTime).c_str(),
				m_lpRecords[i].GlobalObjectIdSize,
				strings::BinToHexString(m_lpRecords[i].lpGlobalObjectId, true).c_str(),
				szGoid.c_str(),
				m_lpRecords[i].UsernameSize,
				m_lpRecords[i].szUsername.c_str());
		}

		return szTombstoneString;
	}
}