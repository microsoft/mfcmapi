#pragma once
#include "SmartViewParser.h"

class SIDBin : public SmartViewParser
{
private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	std::wstring m_lpSidName;
	std::wstring m_lpSidDomain;
	std::wstring m_lpStringSid;
};