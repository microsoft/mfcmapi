#pragma once
#include "SmartViewParser.h"

class SIDBin : public SmartViewParser
{
private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	wstring m_lpSidName;
	wstring m_lpSidDomain;
	wstring m_lpStringSid;
};