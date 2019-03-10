#include <core/stdafx.h>
#include <core/smartview/TombStone.h>
#include <core/smartview/SmartView.h>

namespace smartview
{
	TombstoneRecord::TombstoneRecord(std::shared_ptr<binaryParser> parser)
	{
		StartTime = parser->Get<DWORD>();
		EndTime = parser->Get<DWORD>();
		GlobalObjectIdSize = parser->Get<DWORD>();
		GlobalObjectId.parse(parser, GlobalObjectIdSize, false);
		UsernameSize = parser->Get<WORD>();
		szUsername.init(parser, UsernameSize);
	}

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
				m_lpRecords.emplace_back(std::make_shared<TombstoneRecord>(m_Parser));
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

		auto i = ULONG{};
		for (const auto& record : m_lpRecords)
		{
			terminateBlock();
			addHeader(L"Record[%1!d!]\r\n", i++);
			addBlock(
				record->StartTime,
				L"StartTime = 0x%1!08X! = %2!ws!\r\n",
				record->StartTime.getData(),
				RTimeToString(record->StartTime).c_str());
			addBlock(
				record->EndTime,
				L"Endtime = 0x%1!08X! = %2!ws!\r\n",
				record->EndTime.getData(),
				RTimeToString(record->EndTime).c_str());
			addBlock(
				record->GlobalObjectIdSize,
				L"GlobalObjectIdSize = 0x%1!08X!\r\n",
				record->GlobalObjectIdSize.getData());
			addBlock(record->GlobalObjectId.getBlock());
			terminateBlock();

			addBlock(record->UsernameSize, L"UsernameSize= 0x%1!04X!\r\n", record->UsernameSize.getData());
			addBlock(record->szUsername, L"szUsername = %1!hs!", record->szUsername.c_str());
		}
	}
} // namespace smartview