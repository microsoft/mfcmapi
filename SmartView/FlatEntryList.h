#pragma once
#include "SmartViewParser.h"
#include "SmartView.h" //TODO: Replace with EntryID parser when available

struct FlatEntryIDStruct
{
	DWORD dwSize;
	EntryIdStruct* lpEntryID;

	size_t JunkDataSize;
	LPBYTE JunkData;
};

class FlatEntryList : public SmartViewParser
{
public:
	FlatEntryList(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~FlatEntryList();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_cEntries;
	DWORD m_cbEntries;
	FlatEntryIDStruct* m_pEntryIDs;
};