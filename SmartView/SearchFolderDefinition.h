#pragma once
#include "SmartViewParser.h"
#include "EntryList.h"

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
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_Version;
	DWORD m_Flags;
	DWORD m_NumericSearch;
	BYTE m_TextSearchLength;
	WORD m_TextSearchLengthExtended;
	wstring m_TextSearch;
	DWORD m_SkipLen1;
	vector<BYTE> m_SkipBytes1;
	DWORD m_DeepSearch;
	BYTE m_FolderList1Length;
	WORD m_FolderList1LengthExtended;
	wstring m_FolderList1;
	DWORD m_FolderList2Length;
	EntryList m_FolderList2;
	DWORD m_AddressCount; // SFST_BINARY
	vector<AddressListEntryStruct> m_Addresses; // SFST_BINARY
	DWORD m_SkipLen2;
	vector<BYTE> m_SkipBytes2;
	wstring m_Restriction; // SFST_MRES
	vector<BYTE> m_AdvancedSearchBytes; // SFST_FILTERSTREAM
	DWORD m_SkipLen3;
	vector<BYTE> m_SkipBytes3;
};