#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/EntryIdStruct.h>

namespace smartview
{
	struct EntryListEntryStruct
	{
		blockT<DWORD> EntryLength;
		blockT<DWORD> EntryLengthPad;
		EntryIdStruct EntryId;

		EntryListEntryStruct(std::shared_ptr<binaryParser> parser);
	};

	class EntryList : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_EntryCount;
		blockT<DWORD> m_Pad;

		std::vector<EntryListEntryStruct> m_Entry;
	};
} // namespace smartview