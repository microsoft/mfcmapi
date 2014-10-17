#pragma once
#include "SmartViewParser.h"

class RecipientRowStream : public SmartViewParser
{
public:
	RecipientRowStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~RecipientRowStream();

	_Check_return_ LPWSTR ToString();

private:
	void Parse();

	DWORD m_cVersion;
	DWORD m_cRowCount;
	LPADRENTRY m_lpAdrEntry;
};