#include "stdafx.h"
#include "TombStone.h"
#include "String.h"
#include "SmartView.h"

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
	m_Parser.GetDWORD(&m_Identifier);
	m_Parser.GetDWORD(&m_HeaderSize);
	m_Parser.GetDWORD(&m_Version);
	m_Parser.GetDWORD(&m_RecordsCount);
	m_Parser.GetDWORD(&m_RecordsSize);

	// Run through the parser once to count the number of flag structs
	auto ulFlagOffset = m_Parser.GetCurrentOffset();
	for (;;)
	{
		// Must have at least 2 bytes left to have another flag
		if (m_Parser.RemainingBytes() < sizeof(DWORD) * 3 + sizeof(WORD)) break;
		DWORD dwData = NULL;
		WORD wData = NULL;
		m_Parser.GetDWORD(&dwData);
		m_Parser.GetDWORD(&dwData);
		m_Parser.GetDWORD(&dwData);
		m_Parser.Advance(dwData);
		m_Parser.GetWORD(&wData);
		m_Parser.Advance(wData);
		m_ActualRecordsCount++;
	}

	// Now we parse for real
	m_Parser.SetCurrentOffset(ulFlagOffset);

	if (m_ActualRecordsCount && m_ActualRecordsCount < _MaxEntriesSmall)
	{
		for (ULONG i = 0; i < m_ActualRecordsCount; i++)
		{
			TombstoneRecord tombstoneRecord;
			m_Parser.GetDWORD(&tombstoneRecord.StartTime);
			m_Parser.GetDWORD(&tombstoneRecord.EndTime);
			m_Parser.GetDWORD(&tombstoneRecord.GlobalObjectIdSize);
			tombstoneRecord.lpGlobalObjectId = m_Parser.GetBYTES(tombstoneRecord.GlobalObjectIdSize, _MaxBytes);
			m_Parser.GetWORD(&tombstoneRecord.UsernameSize);
			tombstoneRecord.szUsername = m_Parser.GetStringA(tombstoneRecord.UsernameSize);
			m_lpRecords.push_back(tombstoneRecord);
		}
	}
}

_Check_return_ wstring TombStone::ToStringInternal()
{
	wstring szTombstoneString;

	szTombstoneString = formatmessage(IDS_TOMBSTONEHEADER,
		m_Identifier,
		m_HeaderSize,
		m_Version,
		m_RecordsCount,
		m_ActualRecordsCount,
		m_RecordsSize);

	for (ULONG i = 0; i < m_lpRecords.size(); i++)
	{
		SBinary sBin = { 0 };
		sBin.cb = static_cast<ULONG>(m_lpRecords[i].lpGlobalObjectId.size());
		sBin.lpb = m_lpRecords[i].lpGlobalObjectId.data();
		auto szGoid = InterpretBinaryAsString(sBin, IDS_STGLOBALOBJECTID, nullptr);

		szTombstoneString += formatmessage(IDS_TOMBSTONERECORD,
			i,
			m_lpRecords[i].StartTime, RTimeToString(m_lpRecords[i].StartTime).c_str(),
			m_lpRecords[i].EndTime, RTimeToString(m_lpRecords[i].EndTime).c_str(),
			m_lpRecords[i].GlobalObjectIdSize,
			BinToHexString(m_lpRecords[i].lpGlobalObjectId, true).c_str(),
			szGoid.c_str(),
			m_lpRecords[i].UsernameSize,
			m_lpRecords[i].szUsername.c_str());
	}

	return szTombstoneString;
}