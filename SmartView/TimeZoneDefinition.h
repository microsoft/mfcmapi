#pragma once
#include "SmartViewParser.h"

// TZRule
// =====================
//   This structure represents both a description when a daylight.
//   savings shift occurs, and in addition, the year in which that
//   timezone rule came into effect.
//
struct TZRule
{
	BYTE bMajorVersion;
	BYTE bMinorVersion;
	WORD wReserved;
	WORD wTZRuleFlags;
	WORD wYear;
	BYTE X[14];
	DWORD lBias; // offset from GMT
	DWORD lStandardBias; // offset from bias during standard time
	DWORD lDaylightBias; // offset from bias during daylight time
	SYSTEMTIME stStandardDate; // time to switch to standard time
	SYSTEMTIME stDaylightDate; // time to switch to daylight time
};

// TimeZoneDefinition
// =====================
//   This represents an entire timezone including all historical, current
//   and future timezone shift rules for daylight savings time, etc.
class TimeZoneDefinition : public SmartViewParser
{
public:
	TimeZoneDefinition(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~TimeZoneDefinition();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	BYTE m_bMajorVersion;
	BYTE m_bMinorVersion;
	WORD m_cbHeader;
	WORD m_wReserved;
	WORD m_cchKeyName;
	LPWSTR m_szKeyName;
	WORD m_cRules;
	TZRule* m_lpTZRule;
};