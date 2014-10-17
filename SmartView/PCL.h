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

	_Check_return_ LPWSTR ToString();

private:
	void Parse();

	DWORD m_cXID;
	SizedXID* m_lpXID;
};