#pragma once
#include "SmartViewParser.h"

struct TombstoneRecord
{
	DWORD StartTime;
	DWORD EndTime;
	DWORD GlobalObjectIdSize;
	vector<BYTE> lpGlobalObjectId;
	WORD UsernameSize;
	string szUsername;
};

class TombStone : public SmartViewParser
{
public:
	TombStone();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_Identifier;
	DWORD m_HeaderSize;
	DWORD m_Version;
	DWORD m_RecordsCount;
	DWORD m_ActualRecordsCount; // computed based on state, not read value
	DWORD m_RecordsSize;
	vector<TombstoneRecord> m_lpRecords;
};