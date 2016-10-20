#pragma once
#include "SmartViewParser.h"

class SIDBin : public SmartViewParser
{
public:
	SIDBin();
	~SIDBin();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	LPWSTR m_lpSidName;
	LPWSTR m_lpSidDomain;
	wstring m_lpStringSid;
};