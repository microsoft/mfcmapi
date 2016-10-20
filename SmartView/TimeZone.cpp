#include "stdafx.h"
#include "TimeZone.h"

TimeZone::TimeZone(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	Init(cbBin, lpBin);
	m_lBias = 0;
	m_lStandardBias = 0;
	m_lDaylightBias = 0;
	m_wStandardYear = 0;
	m_stStandardDate = { 0 };
	m_wDaylightDate = 0;
	m_stDaylightDate = { 0 };
}

void TimeZone::Parse()
{
	m_Parser.GetDWORD(&m_lBias);
	m_Parser.GetDWORD(&m_lStandardBias);
	m_Parser.GetDWORD(&m_lDaylightBias);
	m_Parser.GetWORD(&m_wStandardYear);
	m_Parser.GetWORD(&m_stStandardDate.wYear);
	m_Parser.GetWORD(&m_stStandardDate.wMonth);
	m_Parser.GetWORD(&m_stStandardDate.wDayOfWeek);
	m_Parser.GetWORD(&m_stStandardDate.wDay);
	m_Parser.GetWORD(&m_stStandardDate.wHour);
	m_Parser.GetWORD(&m_stStandardDate.wMinute);
	m_Parser.GetWORD(&m_stStandardDate.wSecond);
	m_Parser.GetWORD(&m_stStandardDate.wMilliseconds);
	m_Parser.GetWORD(&m_wDaylightDate);
	m_Parser.GetWORD(&m_stDaylightDate.wYear);
	m_Parser.GetWORD(&m_stDaylightDate.wMonth);
	m_Parser.GetWORD(&m_stDaylightDate.wDayOfWeek);
	m_Parser.GetWORD(&m_stDaylightDate.wDay);
	m_Parser.GetWORD(&m_stDaylightDate.wHour);
	m_Parser.GetWORD(&m_stDaylightDate.wMinute);
	m_Parser.GetWORD(&m_stDaylightDate.wSecond);
	m_Parser.GetWORD(&m_stDaylightDate.wMilliseconds);
}

_Check_return_ wstring TimeZone::ToStringInternal()
{
	wstring szTimeZone;

	szTimeZone = formatmessage(IDS_TIMEZONE,
		m_lBias,
		m_lStandardBias,
		m_lDaylightBias,
		m_wStandardYear,
		m_stStandardDate.wYear,
		m_stStandardDate.wMonth,
		m_stStandardDate.wDayOfWeek,
		m_stStandardDate.wDay,
		m_stStandardDate.wHour,
		m_stStandardDate.wMinute,
		m_stStandardDate.wSecond,
		m_stStandardDate.wMilliseconds,
		m_wDaylightDate,
		m_stDaylightDate.wYear,
		m_stDaylightDate.wMonth,
		m_stDaylightDate.wDayOfWeek,
		m_stDaylightDate.wDay,
		m_stDaylightDate.wHour,
		m_stDaylightDate.wMinute,
		m_stDaylightDate.wSecond,
		m_stDaylightDate.wMilliseconds);

	return szTimeZone;
}