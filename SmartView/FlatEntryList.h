#pragma once
#include "SmartViewParser.h"
#include "EntryIdStruct.h"

struct FlatEntryIDStruct
{
	DWORD dwSize;
	EntryIdStruct* lpEntryID;

	size_t JunkDataSize; // TODO: Kill this junk data
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