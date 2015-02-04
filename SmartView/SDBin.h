#pragma once
#include "SmartViewParser.h"

class SDBin : public SmartViewParser
{
public:
	SDBin(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _In_opt_ LPMAPIPROP lpMAPIProp, bool bFB);
	~SDBin();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	LPMAPIPROP m_lpMAPIProp;
	bool m_bFB;
};