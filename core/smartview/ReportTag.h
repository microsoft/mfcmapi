#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOMSG].pdf
	class ReportTag : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;
		void addEID(const std::wstring& label, const blockT<ULONG>& cb, std::shared_ptr<blockBytes>& eid);

		std::shared_ptr<blockBytes> m_Cookie = emptyBB(); // 8 characters + NULL terminator
		blockT<DWORD> m_Version;
		blockT<ULONG> m_cbStoreEntryID;
		std::shared_ptr<blockBytes> m_lpStoreEntryID = emptyBB();
		blockT<ULONG> m_cbFolderEntryID;
		std::shared_ptr<blockBytes> m_lpFolderEntryID = emptyBB();
		blockT<ULONG> m_cbMessageEntryID;
		std::shared_ptr<blockBytes> m_lpMessageEntryID = emptyBB();
		blockT<ULONG> m_cbSearchFolderEntryID;
		std::shared_ptr<blockBytes> m_lpSearchFolderEntryID = emptyBB();
		blockT<ULONG> m_cbMessageSearchKey;
		std::shared_ptr<blockBytes> m_lpMessageSearchKey = emptyBB();
		blockT<ULONG> m_cchAnsiText;
		std::shared_ptr<blockStringA> m_lpszAnsiText = emptySA();
	};
} // namespace smartview