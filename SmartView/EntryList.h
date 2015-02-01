#pragma once
#include "SmartViewParser.h"
#include "EntryIdStruct.h"

struct EntryListEntryStruct
{
	DWORD EntryLength;
	DWORD EntryLengthPad;
	EntryIdStruct* EntryId;
};

class EntryList : public SmartViewParser
{
public:
	EntryList(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~EntryList();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_EntryCount;
	DWORD m_Pad;

	EntryListEntryStruct* m_Entry;
};