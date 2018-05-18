#pragma once
#include "SmartViewParser.h"

// [MS-OXOTASK].pdf
struct TaskAssigner
{
	DWORD cbAssigner;
	ULONG cbEntryID;
	vector<BYTE> lpEntryID;
	std::string szDisplayName;
	std::wstring wzDisplayName;
	vector<BYTE> JunkData;
};

class TaskAssigners : public SmartViewParser
{
public:
	TaskAssigners();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	DWORD m_cAssigners;
	vector<TaskAssigner> m_lpTaskAssigners;
};