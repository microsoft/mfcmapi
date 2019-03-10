#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	// [MS-OXOMSG].pdf
	class ReportTag : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;
		void addEID(const std::wstring& label, const blockT<ULONG>& cb, const blockBytes& eid);

		blockBytes m_Cookie; // 8 characters + NULL terminator
		blockT<DWORD> m_Version;
		blockT<ULONG> m_cbStoreEntryID;
		blockBytes m_lpStoreEntryID;
		blockT<ULONG> m_cbFolderEntryID;
		blockBytes m_lpFolderEntryID;
		blockT<ULONG> m_cbMessageEntryID;
		blockBytes m_lpMessageEntryID;
		blockT<ULONG> m_cbSearchFolderEntryID;
		blockBytes m_lpSearchFolderEntryID;
		blockT<ULONG> m_cbMessageSearchKey;
		blockBytes m_lpMessageSearchKey;
		blockT<ULONG> m_cchAnsiText;
		blockStringA m_lpszAnsiText;
	};
} // namespace smartview