#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/EntryIdStruct.h>

namespace smartview
{
	struct EntryListEntryStruct
	{
		DWORD EntryLength;
		DWORD EntryLengthPad;
		EntryIdStruct EntryId;
	};

	class EntryList : public SmartViewParser
	{
	public:
		EntryList();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		DWORD m_EntryCount;
		DWORD m_Pad;

		std::vector<EntryListEntryStruct> m_Entry;
	};
}