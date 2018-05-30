#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/EntryList.h>

namespace smartview
{
	struct AddressListEntryStruct
	{
		DWORD PropertyCount;
		DWORD Pad;
		LPSPropValue Props;
	};

	class SearchFolderDefinition : public SmartViewParser
	{
	public:
		SearchFolderDefinition();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		DWORD m_Version;
		DWORD m_Flags;
		DWORD m_NumericSearch;
		BYTE m_TextSearchLength;
		WORD m_TextSearchLengthExtended;
		std::wstring m_TextSearch;
		DWORD m_SkipLen1;
		std::vector<BYTE> m_SkipBytes1;
		DWORD m_DeepSearch;
		BYTE m_FolderList1Length;
		WORD m_FolderList1LengthExtended;
		std::wstring m_FolderList1;
		DWORD m_FolderList2Length;
		EntryList m_FolderList2;
		DWORD m_AddressCount; // SFST_BINARY
		std::vector<AddressListEntryStruct> m_Addresses; // SFST_BINARY
		DWORD m_SkipLen2;
		std::vector<BYTE> m_SkipBytes2;
		std::wstring m_Restriction; // SFST_MRES
		std::vector<BYTE> m_AdvancedSearchBytes; // SFST_FILTERSTREAM
		DWORD m_SkipLen3;
		std::vector<BYTE> m_SkipBytes3;
	};
}