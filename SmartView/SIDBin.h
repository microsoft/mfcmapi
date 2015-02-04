#pragma once
#include "SmartViewParser.h"

class SIDBin : public SmartViewParser
{
public:
	SIDBin(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~SIDBin();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	LPTSTR m_lpSidName;
	LPTSTR m_lpSidDomain;
	wstring m_lpStringSid;
};