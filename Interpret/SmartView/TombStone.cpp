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
					m_Parser.GetBlockBYTES(tombstoneRecord.GlobalObjectIdSize, _MaxBytes);
				tombstoneRecord.UsernameSize = m_Parser.GetBlock<WORD>();
				tombstoneRecord.szUsername = m_Parser.GetBlockStringA(tombstoneRecord.UsernameSize);
				m_lpRecords.push_back(tombstoneRecord);
			}
		}
	}

	void TombStone::ParseBlocks()
	{
		addHeader(L"Tombstone:\r\n");
		addBlock(m_Identifier, L"Identifier = 0x%1!08X!\r\n", m_Identifier.getData());
		addBlock(m_HeaderSize, L"HeaderSize = 0x%1!08X!\r\n", m_HeaderSize.getData());
		addBlock(m_Version, L"Version = 0x%1!08X!\r\n", m_Version.getData());
		addBlock(m_RecordsCount, L"RecordsCount = 0x%1!08X!\r\n", m_RecordsCount.getData());
		addHeader(L"ActualRecordsCount (computed) = 0x%1!08X!\r\n", m_ActualRecordsCount);
		addBlock(m_RecordsSize, L"RecordsSize = 0x%1!08X!", m_RecordsSize.getData());

		for (ULONG i = 0; i < m_lpRecords.size(); i++)
		{
			auto szGoid = InterpretBinaryAsString(
				SBinary{static_cast<ULONG>(m_lpRecords[i].lpGlobalObjectId.getData().size()),
						m_lpRecords[i].lpGlobalObjectId.getData().data()},
				IDS_STGLOBALOBJECTID,
				nullptr);

			addLine();
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
			addBlock(
				m_lpRecords[i].lpGlobalObjectId,
				L"GlobalObjectId = %1!ws!\r\n",
				strings::BinToHexString(m_lpRecords[i].lpGlobalObjectId.getData(), true).c_str());
			addBlock(m_lpRecords[i].lpGlobalObjectId, L"%1!ws!\r\n", szGoid.c_str());
			addBlock(
				m_lpRecords[i].UsernameSize, L"UsernameSize= 0x%1!04X!\r\n", m_lpRecords[i].UsernameSize.getData());
			addBlock(m_lpRecords[i].szUsername, L"szUsername = %1!hs!", m_lpRecords[i].szUsername.c_str());
		}
	}
}