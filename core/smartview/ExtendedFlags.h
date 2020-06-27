#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class ExtendedFlag : public block
	{
	public:
		bool bBadData{};

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> Id = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Cb = emptyT<BYTE>();
		std::shared_ptr<blockT<DWORD>> ExtendedFlags = emptyT<DWORD>();
		std::shared_ptr<blockT<GUID>> SearchFolderID = emptyT<GUID>();
		std::shared_ptr<blockT<DWORD>> SearchFolderTag = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ToDoFolderVersion = emptyT<DWORD>();
		std::shared_ptr<blockBytes> lpUnknownData = emptyBB();
	};

	class ExtendedFlags : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::vector<std::shared_ptr<ExtendedFlag>> m_pefExtendedFlags;
	};
} // namespace smartview