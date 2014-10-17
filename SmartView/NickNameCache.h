#pragma once
#include "SmartViewParser.h"

class NickNameCache : public SmartViewParser
{
public:
	NickNameCache(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~NickNameCache();

	_Check_return_ LPWSTR ToString();

private:
	void Parse();

	BYTE m_Metadata1[4];
	ULONG m_ulMajorVersion;
	ULONG m_ulMinorVersion;
	DWORD m_cRowCount;
	LPSRow m_lpRows;
	ULONG m_cbEI;
	LPBYTE m_lpbEI;
	BYTE m_Metadata2[8];
};