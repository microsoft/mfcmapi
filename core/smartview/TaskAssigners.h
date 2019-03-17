#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOTASK].pdf
	struct TaskAssigner
	{
		blockT<DWORD> cbAssigner{};
		blockT<ULONG> cbEntryID{};
		std::shared_ptr<blockBytes> lpEntryID = emptyBB();
		std::shared_ptr<blockStringA> szDisplayName = emptySA();
		std::shared_ptr<blockStringW> wzDisplayName = emptySW();
		std::shared_ptr<blockBytes> JunkData = emptyBB();

		TaskAssigner(std::shared_ptr<binaryParser>& parser);
	};

	class TaskAssigners : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_cAssigners;
		std::vector<std::shared_ptr<TaskAssigner>> m_lpTaskAssigners;
	};
} // namespace smartview