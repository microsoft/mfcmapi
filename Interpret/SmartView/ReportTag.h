#pragma once
#include "SmartViewParser.h"

// [MS-OXOMSG].pdf
class ReportTag : public SmartViewParser
{
public:
	ReportTag();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	std::vector<BYTE> m_Cookie; // 8 characters + NULL terminator
	DWORD m_Version;
	ULONG m_cbStoreEntryID;
	std::vector<BYTE> m_lpStoreEntryID;
	ULONG m_cbFolderEntryID;
	std::vector<BYTE> m_lpFolderEntryID;
	ULONG m_cbMessageEntryID;
	std::vector<BYTE> m_lpMessageEntryID;
	ULONG m_cbSearchFolderEntryID;
	std::vector<BYTE> m_lpSearchFolderEntryID;
	ULONG m_cbMessageSearchKey;
	std::vector<BYTE> m_lpMessageSearchKey;
	ULONG m_cchAnsiText;
	std::string m_lpszAnsiText;
};