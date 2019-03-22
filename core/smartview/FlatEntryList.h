#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct FlatEntryID
	{
		std::shared_ptr<blockT<DWORD>> dwSize = emptyT<DWORD>();
		EntryIdStruct lpEntryID;

		std::shared_ptr<blockBytes> padding = emptyBB();

		FlatEntryID(std::shared_ptr<binaryParser> parser);
	};

	class FlatEntryList : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_cEntries = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_cbEntries = emptyT<DWORD>();
		std::vector<std::shared_ptr<FlatEntryID>> m_pEntryIDs;
	};
} // namespace smartview