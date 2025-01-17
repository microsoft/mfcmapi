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

		std::shared_ptr<blockBytes> m_Cookie = emptyBB(); // 8 characters + NULL terminator
		std::shared_ptr<blockT<DWORD>> m_Version = emptyT<DWORD>();
		std::shared_ptr<blockT<ULONG>> m_cbStoreEntryID = emptyT<ULONG>();
		std::shared_ptr<blockBytes> m_lpStoreEntryID = emptyBB();
		std::shared_ptr<blockT<ULONG>> m_cbFolderEntryID = emptyT<ULONG>();
		std::shared_ptr<blockBytes> m_lpFolderEntryID = emptyBB();
		std::shared_ptr<blockT<ULONG>> m_cbMessageEntryID = emptyT<ULONG>();
		std::shared_ptr<blockBytes> m_lpMessageEntryID = emptyBB();
		std::shared_ptr<blockT<ULONG>> m_cbSearchFolderEntryID = emptyT<ULONG>();
		std::shared_ptr<blockBytes> m_lpSearchFolderEntryID = emptyBB();
		std::shared_ptr<blockT<ULONG>> m_cbMessageSearchKey = emptyT<ULONG>();
		std::shared_ptr<blockBytes> m_lpMessageSearchKey = emptyBB();
		std::shared_ptr<blockT<ULONG>> m_cchAnsiText = emptyT<ULONG>();
		std::shared_ptr<blockStringA> m_lpszAnsiText = emptySA();
	};
} // namespace smartview