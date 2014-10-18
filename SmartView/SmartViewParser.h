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

	virtual _Check_return_ LPWSTR ToString() = 0;

protected:
	_Check_return_ wstring JunkDataToString();
	_Check_return_ wstring RTimeToString(DWORD rTime);

	ULONG m_cbBin;
	LPBYTE m_lpBin;

	CBinaryParser m_Parser;
};