#pragma once
#include "SmartViewParser.h"

class GlobalObjectId : public SmartViewParser
{
public:
	GlobalObjectId(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~GlobalObjectId();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	BYTE m_Id[16];
	WORD m_Year;
	BYTE m_Month;
	BYTE m_Day;
	FILETIME m_CreationTime;
	LARGE_INTEGER m_X;
	DWORD m_dwSize;
	LPBYTE m_lpData;
};