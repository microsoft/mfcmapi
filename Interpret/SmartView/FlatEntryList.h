#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/EntryIdStruct.h>

namespace smartview
{
	struct FlatEntryID
	{
		blockT<DWORD> dwSize;
		EntryIdStruct lpEntryID;

		blockBytes JunkData;
	};

	class FlatEntryList : public SmartViewParser
	{
	public:
		FlatEntryList();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		blockT<DWORD> m_cEntries;
		blockT<DWORD> m_cbEntries;
		std::vector<FlatEntryID> m_pEntryIDs;
	};
}