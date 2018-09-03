#pragma once
#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/EntryList.h>
#include <Interpret/SmartView/PropertiesStruct.h>

namespace smartview
{
	struct AddressListEntryStruct
	{
		blockT<DWORD> PropertyCount;
		blockT<DWORD> Pad;
		PropertiesStruct Props;
	};

	class SearchFolderDefinition : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_Version;
		blockT<DWORD> m_Flags;
		blockT<DWORD> m_NumericSearch;
		blockT<BYTE> m_TextSearchLength;
		blockT<WORD> m_TextSearchLengthExtended;
		blockStringW m_TextSearch;
		blockT<DWORD> m_SkipLen1;
		blockBytes m_SkipBytes1;
		blockT<DWORD> m_DeepSearch;
		blockT<BYTE> m_FolderList1Length;
		blockT<WORD> m_FolderList1LengthExtended;
		blockStringW m_FolderList1;
		blockT<DWORD> m_FolderList2Length;
		EntryList m_FolderList2;
		blockT<DWORD> m_AddressCount; // SFST_BINARY
		std::vector<AddressListEntryStruct> m_Addresses; // SFST_BINARY
		blockT<DWORD> m_SkipLen2;
		blockBytes m_SkipBytes2;
		blockStringW m_Restriction; // SFST_MRES
		blockBytes m_AdvancedSearchBytes; // SFST_FILTERSTREAM
		blockT<DWORD> m_SkipLen3;
		blockBytes m_SkipBytes3;
	};
}