#pragma once
#include "SmartViewParser.h"

class RecipientRowStream : public SmartViewParser
{
public:
	RecipientRowStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~RecipientRowStream();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_cVersion;
	DWORD m_cRowCount;
	LPADRENTRY m_lpAdrEntry;
};