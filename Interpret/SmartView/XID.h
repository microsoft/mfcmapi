#pragma once
#include "SmartViewParser.h"

class XID : public SmartViewParser
{
public:
	XID();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	GUID m_NamespaceGuid;
	size_t m_cbLocalId;
	vector<BYTE> m_LocalID;
};