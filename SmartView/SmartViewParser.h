#pragma once
#include <string>
using namespace std;
#include "BinaryParser.h"

#define _MaxBytes 0xFFFF
#define _MaxDepth 50
#define _MaxEID 500
#define _MaxEntriesSmall 500
#define _MaxEntriesLarge 1000
#define _MaxEntriesExtraLarge 1500
#define _MaxEntriesEnormous 10000

class SmartViewParser;
typedef SmartViewParser FAR* LPSMARTVIEWPARSER;

class SmartViewParser
{
public:
	SmartViewParser(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	virtual ~SmartViewParser();

	_Check_return_ LPWSTR ToString();

protected:
	_Check_return_ wstring RTimeToString(DWORD rTime);
	_Check_return_ LPSPropValue BinToSPropValue(DWORD dwPropCount, bool bStringPropsExcludeLength);

	CBinaryParser m_Parser;

private:
	virtual void Parse() = 0;
	virtual _Check_return_ wstring ToStringInternal() = 0;
	_Check_return_ wstring JunkDataToString();

	ULONG m_cbBin;
	LPBYTE m_lpBin;
};