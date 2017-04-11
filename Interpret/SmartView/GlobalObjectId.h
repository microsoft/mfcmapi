#pragma once
#include "SmartViewParser.h"

class GlobalObjectId : public SmartViewParser
{
public:
	GlobalObjectId();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	vector<BYTE> m_Id; // 16 bytes
	WORD m_Year;
	BYTE m_Month;
	BYTE m_Day;
	FILETIME m_CreationTime;
	LARGE_INTEGER m_X;
	DWORD m_dwSize;
	vector<BYTE> m_lpData;
};