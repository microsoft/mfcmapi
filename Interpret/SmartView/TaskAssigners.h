#pragma once
#include "SmartViewParser.h"

namespace smartview
{
	// [MS-OXOTASK].pdf
	struct TaskAssigner
	{
		DWORD cbAssigner;
		ULONG cbEntryID;
		std::vector<BYTE> lpEntryID;
		std::string szDisplayName;
		std::wstring wzDisplayName;
		std::vector<BYTE> JunkData;
	};

	class TaskAssigners : public SmartViewParser
	{
	public:
		TaskAssigners();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		DWORD m_cAssigners;
		std::vector<TaskAssigner> m_lpTaskAssigners;
	};
}