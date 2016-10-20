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
	~SearchFolderDefinition();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_Version;
	DWORD m_Flags;
	DWORD m_NumericSearch;
	BYTE m_TextSearchLength;
	WORD m_TextSearchLengthExtended;
	LPWSTR m_TextSearch;
	DWORD m_SkipLen1;
	LPBYTE m_SkipBytes1;
	DWORD m_DeepSearch;
	BYTE m_FolderList1Length;
	WORD m_FolderList1LengthExtended;
	LPWSTR m_FolderList1;
	DWORD m_FolderList2Length;
	EntryList* m_FolderList2;
	DWORD m_AddressCount; // SFST_BINARY
	AddressListEntryStruct* m_Addresses; // SFST_BINARY
	DWORD m_SkipLen2;
	LPBYTE m_SkipBytes2;
	wstring m_Restriction; // SFST_MRES
	DWORD m_AdvancedSearchLen; // SFST_FILTERSTREAM
	LPBYTE m_AdvancedSearchBytes; // SFST_FILTERSTREAM
	DWORD m_SkipLen3;
	LPBYTE m_SkipBytes3;
};