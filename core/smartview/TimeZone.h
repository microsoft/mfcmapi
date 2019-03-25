#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct SYSTEMTIMEBlock
	{
		std::shared_ptr<blockT<WORD>> wYear = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wMonth = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wDayOfWeek = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wDay = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wHour = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wMinute = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wSecond = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wMilliseconds = emptyT<WORD>();
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

		std::shared_ptr<blockT<DWORD>> m_lBias = emptyT<DWORD>(); // offset from GMT
		std::shared_ptr<blockT<DWORD>> m_lStandardBias = emptyT<DWORD>(); // offset from bias during standard time
		std::shared_ptr<blockT<DWORD>> m_lDaylightBias = emptyT<DWORD>(); // offset from bias during daylight time
		std::shared_ptr<blockT<WORD>> m_wStandardYear = emptyT<WORD>();
		SYSTEMTIMEBlock m_stStandardDate; // time to switch to standard time
		std::shared_ptr<blockT<WORD>> m_wDaylightDate = emptyT<WORD>();
		SYSTEMTIMEBlock m_stDaylightDate; // time to switch to daylight time
	};
} // namespace smartview