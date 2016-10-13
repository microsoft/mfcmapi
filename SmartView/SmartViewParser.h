#pragma once
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
	virtual ~SmartViewParser() = default;

	_Check_return_ wstring ToString();

	void DisableJunkParsing();
	size_t GetCurrentOffset() const;
	void EnsureParsed();

protected:
	_Check_return_ LPSPropValue BinToSPropValue(DWORD dwPropCount, bool bStringPropsExcludeLength);
	_Check_return_ wstring JunkDataToString(const vector<BYTE>& lpJunkData) const;
	_Check_return_ wstring JunkDataToString(size_t cbJunkData, _In_count_(cbJunkData) LPBYTE lpJunkData) const;

	CBinaryParser m_Parser;

private:
	virtual void Parse() = 0;
	virtual _Check_return_ wstring ToStringInternal() = 0;

	bool m_bEnableJunk;
	bool m_bParsed;
	ULONG m_cbBin;
	LPBYTE m_lpBin;
};