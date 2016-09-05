#pragma once
#include <mapidefs.h>

class MVPropData
{
public:
	MVPropData(_In_ LPSPropValue lpProp, ULONG iProp);
	MVPropData(_In_ LPSPropValue lpNewValue);
	_PV m_val; // Stolen from input prop
private:
	string lpszA;
	wstring lpszW;
	vector<BYTE> lpBin;
};
