#pragma once
#include <WinNT.h>
#include "SmartViewParser.h"

struct SizedXID
{
	BYTE XidSize;
	GUID NamespaceGuid;
	DWORD cbLocalId;
	LPBYTE LocalID;
};

class PCL : public SmartViewParser
{
public:
	PCL();
	~PCL();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_cXID;
	SizedXID* m_lpXID;
};