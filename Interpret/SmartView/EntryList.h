#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/EntryIdStruct.h>

namespace smartview
{
	struct EntryListEntryStruct
	{
		blockT<DWORD> EntryLength;
		blockT<DWORD> EntryLengthPad;
		EntryIdStruct EntryId;
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