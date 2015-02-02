#pragma once
#include "SmartViewParser.h"

// [MS-OXOMSG].pdf
class ReportTag : public SmartViewParser
{
public:
	ReportTag(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~ReportTag();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	CHAR m_Cookie[9]; // 8 characters + NULL terminator
	DWORD m_Version;
	ULONG m_cbStoreEntryID;
	LPBYTE m_lpStoreEntryID;
	ULONG m_cbFolderEntryID;
	LPBYTE m_lpFolderEntryID;
	ULONG m_cbMessageEntryID;
	LPBYTE m_lpMessageEntryID;
	ULONG m_cbSearchFolderEntryID;
	LPBYTE m_lpSearchFolderEntryID;
	ULONG m_cbMessageSearchKey;
	LPBYTE m_lpMessageSearchKey;
	ULONG m_cchAnsiText;
	LPSTR m_lpszAnsiText;
};