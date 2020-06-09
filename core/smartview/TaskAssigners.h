#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOTASK].pdf
	struct TaskAssigner
	{
		std::shared_ptr<blockT<DWORD>> cbAssigner = emptyT<DWORD>();
		std::shared_ptr<blockT<ULONG>> cbEntryID = emptyT<ULONG>();
		std::shared_ptr<blockBytes> lpEntryID = emptyBB();
		std::shared_ptr<blockStringA> szDisplayName = emptySA();
		std::shared_ptr<blockStringW> wzDisplayName = emptySW();
		std::shared_ptr<blockBytes> JunkData = emptyBB();

		TaskAssigner(const std::shared_ptr<binaryParser>& parser);
	};

	class TaskAssigners : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_cAssigners = emptyT<DWORD>();
		std::vector<std::shared_ptr<TaskAssigner>> m_lpTaskAssigners;
	};
} // namespace smartview