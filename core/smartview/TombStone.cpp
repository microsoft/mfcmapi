#include <core/stdafx.h>
#include <core/smartview/TombStone.h>
#include <core/smartview/SmartView.h>

namespace smartview
{
	void TombstoneRecord::parse()
	{
		StartTime = blockT<DWORD>::parse(parser);
		EndTime = blockT<DWORD>::parse(parser);
		GlobalObjectIdSize = blockT<DWORD>::parse(parser);
		GlobalObjectId = block::parse<smartview::GlobalObjectId>(parser, *GlobalObjectIdSize, false);
		UsernameSize = blockT<WORD>::parse(parser);
		szUsername = blockStringA::parse(parser, *UsernameSize);
	}

	void TombstoneRecord::parseBlocks()
	{
		addChild(StartTime, L"StartTime = 0x%1!08X! = %2!ws!", StartTime->getData(), RTimeToString(*StartTime).c_str());
		addChild(EndTime, L"Endtime = 0x%1!08X! = %2!ws!", EndTime->getData(), RTimeToString(*EndTime).c_str());
		addChild(GlobalObjectIdSize, L"GlobalObjectIdSize = 0x%1!08X!", GlobalObjectIdSize->getData());
		addChild(GlobalObjectId);
		addChild(UsernameSize, L"UsernameSize= 0x%1!04X!", UsernameSize->getData());
		addChild(szUsername, L"szUsername = %1!hs!", szUsername->c_str());
	}

	void TombStone::parse()
	{
		m_Identifier = blockT<DWORD>::parse(parser);
		m_HeaderSize = blockT<DWORD>::parse(parser);
		m_Version = blockT<DWORD>::parse(parser);
		m_RecordsCount = blockT<DWORD>::parse(parser);
		m_RecordsSize = blockT<DWORD>::parse(parser);

		// Run through the parser once to count the number of flag structs
		const auto ulFlagOffset = parser->getOffset();
		auto actualRecordsCount = 0;
		for (;;)
		{
			// Must have at least 2 bytes left to have another flag
			if (parser->getSize() < sizeof(DWORD) * 3 + sizeof(WORD)) break;
			static_cast<void>(parser->advance(sizeof DWORD));
			static_cast<void>(parser->advance(sizeof DWORD));
			const auto len1 = blockT<DWORD>::parse(parser);
			parser->advance(*len1);
			const auto len2 = blockT<WORD>::parse(parser);
			parser->advance(*len2);
			actualRecordsCount++;
		}

		// Now we parse for real
		parser->setOffset(ulFlagOffset);

		if (actualRecordsCount && actualRecordsCount < _MaxEntriesSmall)
		{
			m_lpRecords.reserve(actualRecordsCount);
			for (auto i = 0; i < actualRecordsCount; i++)
			{
				m_lpRecords.emplace_back(block::parse<TombstoneRecord>(parser, false));
			}
		}
	}

	void TombStone::parseBlocks()
	{
		setText(L"Tombstone");
		addChild(m_Identifier, L"Identifier = 0x%1!08X!", m_Identifier->getData());
		addChild(m_HeaderSize, L"HeaderSize = 0x%1!08X!", m_HeaderSize->getData());
		addChild(m_Version, L"Version = 0x%1!08X!", m_Version->getData());
		addChild(m_RecordsCount, L"RecordsCount = 0x%1!08X!", m_RecordsCount->getData());
		addHeader(L"ActualRecordsCount (computed) = 0x%1!08X!", m_lpRecords.size());
		addChild(m_RecordsSize, L"RecordsSize = 0x%1!08X!", m_RecordsSize->getData());

		auto i = 0;
		for (const auto& record : m_lpRecords)
		{
			addChild(record, L"Record[%1!d!]", i++);
		}
	}
} // namespace smartview