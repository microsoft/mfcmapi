#pragma once
#include "SmartViewParser.h"
#include "SmartView.h"
#include <MAPIDefs.h>

class PropertyStruct : public SmartViewParser
{
public:
	PropertyStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~PropertyStruct();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_PropCount;
	LPSPropValue m_Prop;
};

_Check_return_ wstring PropsToString(DWORD PropCount, LPSPropValue Prop);

// Neuters an array of SPropValues - caller must use delete to delete the SPropValue
void DeleteSPropVal(ULONG cVal, _In_count_(cVal) LPSPropValue lpsPropVal);