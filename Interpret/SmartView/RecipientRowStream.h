#pragma once
#include "SmartViewParser.h"

class RecipientRowStream : public SmartViewParser
{
public:
	RecipientRowStream();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_cVersion;
	DWORD m_cRowCount;
	LPADRENTRY m_lpAdrEntry;
};