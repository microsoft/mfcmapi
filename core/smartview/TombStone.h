#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/GlobalObjectId.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class TombstoneRecord : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> StartTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> EndTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> GlobalObjectIdSize = emptyT<DWORD>();
		std::shared_ptr<GlobalObjectId> GlobalObjectId;
		std::shared_ptr<blockT<WORD>> UsernameSize = emptyT<WORD>();
		std::shared_ptr<blockStringA> szUsername = emptySA();
	};

	class TombStone : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_Identifier = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_HeaderSize = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_Version = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_RecordsCount = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_RecordsSize = emptyT<DWORD>();
		std::vector<std::shared_ptr<TombstoneRecord>> m_lpRecords;
	};
} // namespace smartview