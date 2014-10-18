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
	PCL(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~PCL();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_cXID;
	SizedXID* m_lpXID;
};