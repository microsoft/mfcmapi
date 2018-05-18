#pragma once
#include <mapidefs.h>
#include "Data.h"

class MVPropData :public IData
{
public:
	MVPropData(_In_opt_ LPSPropValue lpProp, ULONG iProp);
	MVPropData(_In_opt_ LPSPropValue lpNewValue);
	_PV m_val;
private:
	std::string m_lpszA;
	std::wstring m_lpszW;
	vector<BYTE> m_lpBin;
};
