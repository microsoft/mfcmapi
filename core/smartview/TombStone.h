#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/GlobalObjectId.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct TombstoneRecord
	{
		std::shared_ptr<blockT<DWORD>> StartTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> EndTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> GlobalObjectIdSize = emptyT<DWORD>();
		GlobalObjectId GlobalObjectId;
		std::shared_ptr<blockT<WORD>> UsernameSize = emptyT<WORD>();
		std::shared_ptr<blockStringA> szUsername = emptySA();

		TombstoneRecord(std::shared_ptr<binaryParser> parser);
	};

	class TombStone : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_Identifier = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_HeaderSize = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_Version = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_RecordsCount = emptyT<DWORD>();
//		DWORD m_ActualRecordsCount{}; // computed based on state, not read value
		std::shared_ptr<blockT<DWORD>> m_RecordsSize = emptyT<DWORD>();
		std::vector<std::shared_ptr<TombstoneRecord>> m_lpRecords;
	};
} // namespace smartview