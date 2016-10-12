#pragma once
#include <mapidefs.h>

class MVPropData
{
public:
	MVPropData(_In_opt_ LPSPropValue lpProp, ULONG iProp);
	MVPropData(_In_opt_ LPSPropValue lpNewValue);
	_PV m_val;
private:
	string m_lpszA;
	wstring m_lpszW;
	vector<BYTE> m_lpBin;
};
