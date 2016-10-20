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
	FlatEntryList();
	~FlatEntryList();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_cEntries;
	DWORD m_cbEntries;
	FlatEntryIDStruct* m_pEntryIDs;
};