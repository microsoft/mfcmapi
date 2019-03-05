#include <core/stdafx.h>
#include <core/smartview/TombStone.h>
#include <core/smartview/SmartView.h>

namespace smartview
{
	void TombStone::Parse()
	{
		m_Identifier = m_Parser->Get<DWORD>();
		m_HeaderSize = m_Parser->Get<DWORD>();
		m_Version = m_Parser->Get<DWORD>();
		m_RecordsCount = m_Parser->Get<DWORD>();
		m_RecordsSize = m_Parser->Get<DWORD>();

		// Run through the parser once to count the number of flag structs
		const auto ulFlagOffset = m_Parser->GetCurrentOffset();
		for (;;)
		{
			// Must have at least 2 bytes left to have another flag
			if (m_Parser->RemainingBytes() < sizeof(DWORD) * 3 + sizeof(WORD)) break;
			(void) m_Parser->Get<DWORD>();
			(void) m_Parser->Get<DWORD>();
			m_Parser->advance(m_Parser->Get<DWORD>());
			m_Parser->advance(m_Parser->Get<WORD>());
			m_ActualRecordsCount++;
		}

		// Now we parse for real
		m_Parser->SetCurrentOffset(ulFlagOffset);

		if (m_ActualRecordsCount && m_ActualRecordsCount < _MaxEntriesSmall)
		{
			m_lpRecords.reserve(m_ActualRecordsCount);
			for (ULONG i = 0; i < m_ActualRecordsCount; i++)
			{
				TombstoneRecord tombstoneRecord;
				tombstoneRecord.StartTime = m_Parser->Get<DWORD>();
				tombstoneRecord.EndTime = m_Parser->Get<DWORD>();
				tombstoneRecord.GlobalObjectIdSize = m_Parser->Get<DWORD>();
				tombstoneRecord.GlobalObjectId.parse(m_Parser, tombstoneRecord.GlobalObjectIdSize, false);
				tombstoneRecord.UsernameSize = m_Parser->Get<WORD>();
				tombstoneRecord.szUsername = m_Parser->GetStringA(tombstoneRecord.UsernameSize);
				m_lpRecords.push_back(tombstoneRecord);
			}
		}
	}

	void TombStone::ParseBlocks()
	{
		setRoot(L"Tombstone:\r\n");
		addBlock(m_Identifier, L"Identifier = 0x%1!08X!\r\n", m_Identifier.getData());
		addBlock(m_HeaderSize, L"HeaderSize = 0x%1!08X!\r\n", m_HeaderSize.getData());
		addBlock(m_Version, L"Version = 0x%1!08X!\r\n", m_Version.getData());
		addBlock(m_RecordsCount, L"RecordsCount = 0x%1!08X!\r\n", m_RecordsCount.getData());
		addHeader(L"ActualRecordsCount (computed) = 0x%1!08X!\r\n", m_ActualRecordsCount);
		addBlock(m_RecordsSize, L"RecordsSize = 0x%1!08X!", m_RecordsSize.getData());

		for (ULONG i = 0; i < m_lpRecords.size(); i++)
		{
			terminateBlock();
			addHeader(L"Record[%1!d!]\r\n", i);
			addBlock(
				m_lpRecords[i].StartTime,
				L"StartTime = 0x%1!08X! = %2!ws!\r\n",
				m_lpRecords[i].StartTime.getData(),
				RTimeToString(m_lpRecords[i].StartTime).c_str());
			addBlock(
				m_lpRecords[i].EndTime,
				L"Endtime = 0x%1!08X! = %2!ws!\r\n",
				m_lpRecords[i].EndTime.getData(),
				RTimeToString(m_lpRecords[i].EndTime).c_str());
			addBlock(
				m_lpRecords[i].GlobalObjectIdSize,
				L"GlobalObjectIdSize = 0x%1!08X!\r\n",
				m_lpRecords[i].GlobalObjectIdSize.getData());
			addBlock(m_lpRecords[i].GlobalObjectId.getBlock());
			terminateBlock();

			addBlock(
				m_lpRecords[i].UsernameSize, L"UsernameSize= 0x%1!04X!\r\n", m_lpRecords[i].UsernameSize.getData());
			addBlock(m_lpRecords[i].szUsername, L"szUsername = %1!hs!", m_lpRecords[i].szUsername.c_str());
		}
	}
} // namespace smartview