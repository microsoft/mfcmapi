#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	// [MS-OXOTASK].pdf
	struct TaskAssigner
	{
		blockT<DWORD> cbAssigner{};
		blockT<ULONG> cbEntryID{};
		blockBytes lpEntryID;
		blockStringA szDisplayName;
		blockStringW wzDisplayName;
		blockBytes JunkData;
	};

	class TaskAssigners : public SmartViewParser
	{
	public:
		TaskAssigners();

	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_cAssigners;
		std::vector<TaskAssigner> m_lpTaskAssigners;
	};
}