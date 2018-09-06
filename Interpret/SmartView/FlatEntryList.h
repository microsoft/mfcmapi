#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/EntryIdStruct.h>

namespace smartview
{
	// FlatEntryID
	// https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/flatentry
	struct FlatEntryID
	{
		DWORD dwSize;
		EntryIdStruct lpEntryID;

		std::vector<BYTE> m_Padding; // 0-3 bytes
	};

	// FlatEntryList
	// https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/flatentrylist
	class FlatEntryList : public SmartViewParser
	{
	public:
		FlatEntryList();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		DWORD m_cEntries{};
		DWORD m_cbEntries{};
		std::vector<FlatEntryID> m_pEntryIDs;
	};
}