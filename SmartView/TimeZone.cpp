#include "stdafx.h"
#include "TimeZone.h"

TimeZone::TimeZone()
{
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
	m_lBias = m_Parser.Get<DWORD>();
	m_lStandardBias = m_Parser.Get<DWORD>();
	m_lDaylightBias = m_Parser.Get<DWORD>();
	m_wStandardYear = m_Parser.Get<WORD>();
	m_stStandardDate.wYear = m_Parser.Get<WORD>();
	m_stStandardDate.wMonth = m_Parser.Get<WORD>();
	m_stStandardDate.wDayOfWeek = m_Parser.Get<WORD>();
	m_stStandardDate.wDay = m_Parser.Get<WORD>();
	m_stStandardDate.wHour = m_Parser.Get<WORD>();
	m_stStandardDate.wMinute = m_Parser.Get<WORD>();
	m_stStandardDate.wSecond = m_Parser.Get<WORD>();
	m_stStandardDate.wMilliseconds = m_Parser.Get<WORD>();
	m_wDaylightDate = m_Parser.Get<WORD>();
	m_stDaylightDate.wYear = m_Parser.Get<WORD>();
	m_stDaylightDate.wMonth = m_Parser.Get<WORD>();
	m_stDaylightDate.wDayOfWeek = m_Parser.Get<WORD>();
	m_stDaylightDate.wDay = m_Parser.Get<WORD>();
	m_stDaylightDate.wHour = m_Parser.Get<WORD>();
	m_stDaylightDate.wMinute = m_Parser.Get<WORD>();
	m_stDaylightDate.wSecond = m_Parser.Get<WORD>();
	m_stDaylightDate.wMilliseconds = m_Parser.Get<WORD>();
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