#pragma once
#include "SmartViewParser.h"

// [MS-OXOTASK].pdf
struct TaskAssigner
{
	DWORD cbAssigner;
	ULONG cbEntryID;
	LPBYTE lpEntryID;
	LPSTR szDisplayName;
	LPWSTR wzDisplayName;
	size_t JunkDataSize;
	LPBYTE JunkData; // TODO: Kill this
};

class TaskAssigners : public SmartViewParser
{
public:
	TaskAssigners();
	~TaskAssigners();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_cAssigners;
	TaskAssigner* m_lpTaskAssigners;
};