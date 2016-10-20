#pragma once
#include "SmartViewParser.h"

class XID : public SmartViewParser
{
public:
	XID();
	~XID();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	GUID m_NamespaceGuid;
	size_t m_cbLocalId;
	LPBYTE m_LocalID;
};