#include "stdafx.h"
#include "..\stdafx.h"
#include "TombStone.h"
#include "BinaryParser.h"
#include "..\String.h"
#include "..\ParseProperty.h"
#include "SmartView.h"

TombStone::TombStone(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_Identifier = 0;
	m_HeaderSize = 0;
	m_Version = 0;
	m_RecordsCount = 0;
	m_ActualRecordsCount = 0;
	m_RecordsSize = 0;
	m_lpRecords = NULL;
}

TombStone::~TombStone()
{
	if (m_RecordsCount && m_lpRecords)
	{
		ULONG i = 0;

		for (i = 0; i < m_ActualRecordsCount; i++)
		{
			delete[] m_lpRecords[i].lpGlobalObjectId;
			delete[] m_lpRecords[i].szUsername;
		}

		delete[] m_lpRecords;
	}
}

void TombStone::Parse()
{
	if (!m_lpBin) return;

	CBinaryParser Parser(m_cbBin, m_lpBin);

	Parser.GetDWORD(&m_Identifier);
	Parser.GetDWORD(&m_HeaderSize);
	Parser.GetDWORD(&m_Version);
	Parser.GetDWORD(&m_RecordsCount);
	Parser.GetDWORD(&m_RecordsSize);

	// Run through the parser once to count the number of flag structs
	CBinaryParser Parser2(m_cbBin, m_lpBin);
	Parser2.SetCurrentOffset(Parser.GetCurrentOffset());
	for (;;)
	{
		// Must have at least 2 bytes left to have another flag
		if (Parser2.RemainingBytes() < sizeof(DWORD)* 3 + sizeof(WORD)) break;
		DWORD dwData = NULL;
		WORD wData = NULL;
		Parser2.GetDWORD(&dwData);
		Parser2.GetDWORD(&dwData);
		Parser2.GetDWORD(&dwData);
		Parser2.Advance(dwData);
		Parser2.GetWORD(&wData);
		Parser2.Advance(wData);
		m_ActualRecordsCount++;
	}

	if (m_ActualRecordsCount && m_ActualRecordsCount < _MaxEntriesSmall)
		m_lpRecords = new TombstoneRecord[m_ActualRecordsCount];
	if (m_lpRecords)
	{
		memset(m_lpRecords, 0, sizeof(TombstoneRecord)*m_ActualRecordsCount);
		ULONG i = 0;

		for (i = 0; i < m_ActualRecordsCount; i++)
		{
			Parser.GetDWORD(&m_lpRecords[i].StartTime);
			Parser.GetDWORD(&m_lpRecords[i].EndTime);
			Parser.GetDWORD(&m_lpRecords[i].GlobalObjectIdSize);
			Parser.GetBYTES(m_lpRecords[i].GlobalObjectIdSize, _MaxBytes, &m_lpRecords[i].lpGlobalObjectId);
			Parser.GetWORD(&m_lpRecords[i].UsernameSize);
			Parser.GetStringA(m_lpRecords[i].UsernameSize, &m_lpRecords[i].szUsername);
		}
	}

	m_JunkDataSize = Parser.GetRemainingData(&m_JunkData);
}

_Check_return_ LPWSTR TombStone::ToString()
{
	Parse();

	wstring szTombstoneString;
	wstring szTmp;

	szTombstoneString = formatmessage(IDS_TOMBSTONEHEADER,
		m_Identifier,
		m_HeaderSize,
		m_Version,
		m_RecordsCount,
		m_ActualRecordsCount,
		m_RecordsSize);

	if (m_RecordsCount && m_lpRecords)
	{
		ULONG i = 0;
		for (i = 0; i < m_ActualRecordsCount; i++)
		{
			LPWSTR szGoid = NULL;
			SBinary sBin = { 0 };
			sBin.cb = m_lpRecords[i].GlobalObjectIdSize;
			sBin.lpb = m_lpRecords[i].lpGlobalObjectId;
			InterpretBinaryAsString(sBin, IDS_STGLOBALOBJECTID, NULL, NULL, &szGoid);

			szTmp = formatmessage(IDS_TOMBSTONERECORD,
				i,
				m_lpRecords[i].StartTime, RTimeToString(m_lpRecords[i].StartTime).c_str(),
				m_lpRecords[i].EndTime, RTimeToString(m_lpRecords[i].EndTime).c_str(),
				m_lpRecords[i].GlobalObjectIdSize,
				BinToHexString(&sBin, true).c_str(),
				szGoid,
				m_lpRecords[i].UsernameSize,
				m_lpRecords[i].szUsername);
			szTombstoneString += szTmp;

			delete[] szGoid;
		}
	}

	szTombstoneString += JunkDataToString();

	return wstringToLPWSTR(szTombstoneString);
}