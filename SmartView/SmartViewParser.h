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
	SmartViewParser();
	virtual ~SmartViewParser() = default;

	void Init(size_t cbBin, _In_count_(cbBin) LPBYTE lpBin);
	_Check_return_ wstring ToString();

	void DisableJunkParsing();
	size_t GetCurrentOffset() const;
	void EnsureParsed();

protected:
	_Check_return_ wstring JunkDataToString(const vector<BYTE>& lpJunkData) const;
	_Check_return_ wstring JunkDataToString(size_t cbJunkData, _In_count_(cbJunkData) LPBYTE lpJunkData) const;
	_Check_return_ LPSPropValue BinToSPropValue(DWORD dwPropCount, bool bStringPropsExcludeLength);

	// These functions return pointers to memory backed and cleaned up by SmartViewParser
	LPBYTE GetBYTES(size_t cbBytes, size_t cbMaxBytes = -1);
	LPSTR GetStringA(size_t cchChar = -1);
	LPWSTR GetStringW(size_t cchChar = -1);
	LPBYTE Allocate(size_t cbBytes);
	LPBYTE AllocateArray(size_t cArray, size_t cbEntry);

	CBinaryParser m_Parser;

private:
	virtual void Parse() = 0;
	virtual _Check_return_ wstring ToStringInternal() = 0;

	bool m_bEnableJunk;
	bool m_bParsed;

	// We use list instead of vector so our nodes never get reallocated
	std::list<string> m_stringCache;
	std::list<wstring> m_wstringCache;
	std::list<vector<BYTE>> m_binCache;
};