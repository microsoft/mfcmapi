#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOMSG] 2.2.2.22 PidTagReportTag Property
	// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxomsg/b90b2c88-fe59-4fd9-a6a7-ec18833654d6
	class ReportTag : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;
		void addEID(
			const std::wstring& label,
			const std::shared_ptr<blockT<ULONG>>& cb,
			const std::shared_ptr<blockBytes>& eid);

		std::shared_ptr<blockBytes> Cookie = emptyBB(); // 8 characters + NULL terminator
		std::shared_ptr<blockT<DWORD>> Version = emptyT<DWORD>();
		std::shared_ptr<blockT<ULONG>> StoreEntryIdSize = emptyT<ULONG>();
		std::shared_ptr<blockBytes> StoreEntryId = emptyBB();
		std::shared_ptr<blockT<ULONG>> FolderEntryIdSize = emptyT<ULONG>();
		std::shared_ptr<blockBytes> FolderEntryId = emptyBB();
		std::shared_ptr<blockT<ULONG>> MessageEntryIdSize = emptyT<ULONG>();
		std::shared_ptr<blockBytes> MessageEntryId = emptyBB();
		std::shared_ptr<blockT<ULONG>> SearchFolderEntryIdSize = emptyT<ULONG>();
		std::shared_ptr<blockBytes> SearchFolderEntryId = emptyBB();
		std::shared_ptr<blockT<ULONG>> MessageSearchKeySize = emptyT<ULONG>();
		std::shared_ptr<blockBytes> MessageSearchKey = emptyBB();
		std::shared_ptr<blockT<ULONG>> ANSITextSize = emptyT<ULONG>();
		std::shared_ptr<blockStringA> ANSIText = emptySA();
	};
} // namespace smartview