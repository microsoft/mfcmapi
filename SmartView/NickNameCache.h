#pragma once
#include "SmartViewParser.h"

class NickNameCache : public SmartViewParser
{
public:
	NickNameCache();
	~NickNameCache();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	_Check_return_ LPSPropValue NickNameBinToSPropValue(DWORD dwPropCount);

	BYTE m_Metadata1[4];
	ULONG m_ulMajorVersion;
	ULONG m_ulMinorVersion;
	DWORD m_cRowCount;
	LPSRow m_lpRows;
	ULONG m_cbEI;
	LPBYTE m_lpbEI;
	BYTE m_Metadata2[8];
};