#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	// [MS-OXOMSG].pdf
	class ReportTag : public SmartViewParser
	{
	public:
		ReportTag();

	private:
		void Parse() override;
		void ParseBlocks() override;

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
		blockT<std::string> m_lpszAnsiText;
	};
}