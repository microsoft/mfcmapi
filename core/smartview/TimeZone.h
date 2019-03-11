#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct SYSTEMTIMEBlock
	{
		blockT<WORD> wYear;
		blockT<WORD> wMonth;
		blockT<WORD> wDayOfWeek;
		blockT<WORD> wDay;
		blockT<WORD> wHour;
		blockT<WORD> wMinute;
		blockT<WORD> wSecond;
		blockT<WORD> wMilliseconds;
	};

	// TimeZone
	// =====================
	//   This is an individual description that defines when a daylight
	//   savings shift, and the return to standard time occurs, and how
	//   far the shift is.  This is basically the same as
	//   TIME_ZONE_INFORMATION documented in MSDN, except that the strings
	//   describing the names 'daylight' and 'standard' time are omitted.
	class TimeZone : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_lBias; // offset from GMT
		blockT<DWORD> m_lStandardBias; // offset from bias during standard time
		blockT<DWORD> m_lDaylightBias; // offset from bias during daylight time
		blockT<WORD> m_wStandardYear;
		SYSTEMTIMEBlock m_stStandardDate; // time to switch to standard time
		blockT<WORD> m_wDaylightDate;
		SYSTEMTIMEBlock m_stDaylightDate; // time to switch to daylight time
	};
} // namespace smartview