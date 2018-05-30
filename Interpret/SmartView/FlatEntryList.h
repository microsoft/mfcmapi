#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/EntryIdStruct.h>

namespace smartview
{
	struct FlatEntryID
	{
		DWORD dwSize;
		EntryIdStruct lpEntryID;

		std::vector<BYTE> JunkData;
	};

	class FlatEntryList : public SmartViewParser
	{
	public:
		FlatEntryList();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		DWORD m_cEntries;
		DWORD m_cbEntries;
		std::vector<FlatEntryID> m_pEntryIDs;
	};
}