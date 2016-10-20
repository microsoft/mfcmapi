#pragma once
#include <WinNT.h>
#include "SmartViewParser.h"

struct SizedXID
{
	BYTE XidSize;
	GUID NamespaceGuid;
	DWORD cbLocalId;
	vector<BYTE> LocalID;
};

class PCL : public SmartViewParser
{
public:
	PCL();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_cXID;
	vector<SizedXID> m_lpXID;
};