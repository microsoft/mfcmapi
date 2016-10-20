#pragma once
#include "SmartViewParser.h"
#include "EntryIdStruct.h"

struct FlatEntryID
{
	DWORD dwSize;
	EntryIdStruct lpEntryID;

	vector<BYTE> JunkData;
};

class FlatEntryList : public SmartViewParser
{
public:
	FlatEntryList();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_cEntries;
	DWORD m_cbEntries;
	vector<FlatEntryID> m_pEntryIDs;
};