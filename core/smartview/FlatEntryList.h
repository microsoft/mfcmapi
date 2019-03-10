#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	struct FlatEntryID
	{
		blockT<DWORD> dwSize;
		EntryIdStruct lpEntryID;

		blockBytes padding;

		FlatEntryID(std::shared_ptr<binaryParser> parser);
	};

	class FlatEntryList : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_cEntries;
		blockT<DWORD> m_cbEntries;
		std::vector<std::shared_ptr<FlatEntryID>> m_pEntryIDs;
	};
} // namespace smartview