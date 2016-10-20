#pragma once
#include "SmartViewParser.h"

// TimeZone
// =====================
//   This is an individual description that defines when a daylight
//   savings shift, and the return to standard time occurs, and how
//   far the shift is.  This is basically the same as
//   TIME_ZONE_INFORMATION documented in MSDN, except that the strings
//   describing the names 'daylight' and 'standard' time are omitted.
class TimeZone : public SmartViewParser
{
public:
	TimeZone();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_lBias; // offset from GMT
	DWORD m_lStandardBias; // offset from bias during standard time
	DWORD m_lDaylightBias; // offset from bias during daylight time
	WORD m_wStandardYear;
	SYSTEMTIME m_stStandardDate; // time to switch to standard time
	WORD m_wDaylightDate;
	SYSTEMTIME m_stDaylightDate; // time to switch to daylight time
};