#pragma once
#include "SmartViewParser.h"

struct TombstoneRecord
{
	DWORD StartTime;
	DWORD EndTime;
	DWORD GlobalObjectIdSize;
	LPBYTE lpGlobalObjectId;
	WORD UsernameSize;
	LPSTR szUsername;
};

class TombStone : public SmartViewParser
{
public:
	TombStone(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~TombStone();

	_Check_return_ LPWSTR ToString();

private:
	void Parse();

	DWORD m_Identifier;
	DWORD m_HeaderSize;
	DWORD m_Version;
	DWORD m_RecordsCount;
	DWORD m_ActualRecordsCount; // computed based on state, not read value
	DWORD m_RecordsSize;
	TombstoneRecord* m_lpRecords;
};