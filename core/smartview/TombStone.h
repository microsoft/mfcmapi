#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/GlobalObjectId.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct TombstoneRecord
	{
		blockT<DWORD> StartTime;
		blockT<DWORD> EndTime;
		blockT<DWORD> GlobalObjectIdSize;
		GlobalObjectId GlobalObjectId;
		blockT<WORD> UsernameSize;
		blockStringA szUsername;

		TombstoneRecord(std::shared_ptr<binaryParser> parser);
	};

	class TombStone : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_Identifier;
		blockT<DWORD> m_HeaderSize;
		blockT<DWORD> m_Version;
		blockT<DWORD> m_RecordsCount;
		DWORD m_ActualRecordsCount{}; // computed based on state, not read value
		blockT<DWORD> m_RecordsSize;
		std::vector<std::shared_ptr<TombstoneRecord>> m_lpRecords;
	};
} // namespace smartview