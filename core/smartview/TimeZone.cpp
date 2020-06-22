#include <core/stdafx.h>
#include <core/smartview/TimeZone.h>

namespace smartview
{
	void TimeZone::parse()
	{
		m_lBias = blockT<DWORD>::parse(parser);
		m_lStandardBias = blockT<DWORD>::parse(parser);
		m_lDaylightBias = blockT<DWORD>::parse(parser);
		m_wStandardYear = blockT<WORD>::parse(parser);
		m_stStandardDate.wYear = blockT<WORD>::parse(parser);
		m_stStandardDate.wMonth = blockT<WORD>::parse(parser);
		m_stStandardDate.wDayOfWeek = blockT<WORD>::parse(parser);
		m_stStandardDate.wDay = blockT<WORD>::parse(parser);
		m_stStandardDate.wHour = blockT<WORD>::parse(parser);
		m_stStandardDate.wMinute = blockT<WORD>::parse(parser);
		m_stStandardDate.wSecond = blockT<WORD>::parse(parser);
		m_stStandardDate.wMilliseconds = blockT<WORD>::parse(parser);
		m_wDaylightDate = blockT<WORD>::parse(parser);
		m_stDaylightDate.wYear = blockT<WORD>::parse(parser);
		m_stDaylightDate.wMonth = blockT<WORD>::parse(parser);
		m_stDaylightDate.wDayOfWeek = blockT<WORD>::parse(parser);
		m_stDaylightDate.wDay = blockT<WORD>::parse(parser);
		m_stDaylightDate.wHour = blockT<WORD>::parse(parser);
		m_stDaylightDate.wMinute = blockT<WORD>::parse(parser);
		m_stDaylightDate.wSecond = blockT<WORD>::parse(parser);
		m_stDaylightDate.wMilliseconds = blockT<WORD>::parse(parser);
	}

	void TimeZone::parseBlocks()
	{
		setText(L"Time Zone");
		addChild(m_lBias, L"lBias = 0x%1!08X! (%1!d!)", m_lBias->getData());
		addChild(m_lStandardBias, L"lStandardBias = 0x%1!08X! (%1!d!)", m_lStandardBias->getData());
		addChild(m_lDaylightBias, L"lDaylightBias = 0x%1!08X! (%1!d!)", m_lDaylightBias->getData());
		addChild(m_wStandardYear, L"wStandardYear = 0x%1!04X! (%1!d!)", m_wStandardYear->getData());
		auto standard = create(L"stStandardDate");
		addChild(standard);
		standard->addChild(m_stStandardDate.wYear, L"wYear = 0x%1!X! (%1!d!)", m_stStandardDate.wYear->getData());
		standard->addChild(m_stStandardDate.wMonth, L"wMonth = 0x%1!X! (%1!d!)", m_stStandardDate.wMonth->getData());
		standard->addChild(
			m_stStandardDate.wDayOfWeek, L"wDayOfWeek = 0x%1!X! (%1!d!)", m_stStandardDate.wDayOfWeek->getData());
		standard->addChild(m_stStandardDate.wDay, L"wDay = 0x%1!X! (%1!d!)", m_stStandardDate.wDay->getData());
		standard->addChild(m_stStandardDate.wHour, L"wHour = 0x%1!X! (%1!d!)", m_stStandardDate.wHour->getData());
		standard->addChild(m_stStandardDate.wMinute, L"wMinute = 0x%1!X! (%1!d!)", m_stStandardDate.wMinute->getData());
		standard->addChild(m_stStandardDate.wSecond, L"wSecond = 0x%1!X! (%1!d!)", m_stStandardDate.wSecond->getData());
		standard->addChild(
			m_stStandardDate.wMilliseconds,
			L"wMilliseconds = 0x%1!X! (%1!d!)",
			m_stStandardDate.wMilliseconds->getData());
		addChild(m_wDaylightDate, L"wDaylightDate = 0x%1!04X! (%1!d!)", m_wDaylightDate->getData());
		auto daylight = create(L"stDaylightDate");
		addChild(daylight);
		daylight->addChild(m_stDaylightDate.wYear, L"wYear = 0x%1!X! (%1!d!)", m_stDaylightDate.wYear->getData());
		daylight->addChild(m_stDaylightDate.wMonth, L"wMonth = 0x%1!X! (%1!d!)", m_stDaylightDate.wMonth->getData());
		daylight->addChild(
			m_stDaylightDate.wDayOfWeek, L"wDayOfWeek = 0x%1!X! (%1!d!)", m_stDaylightDate.wDayOfWeek->getData());
		daylight->addChild(m_stDaylightDate.wDay, L"wDay = 0x%1!X! (%1!d!)", m_stDaylightDate.wDay->getData());
		daylight->addChild(m_stDaylightDate.wHour, L"wHour = 0x%1!X! (%1!d!)", m_stDaylightDate.wHour->getData());
		daylight->addChild(m_stDaylightDate.wMinute, L"wMinute = 0x%1!X! (%1!d!)", m_stDaylightDate.wMinute->getData());
		daylight->addChild(m_stDaylightDate.wSecond, L"wSecond = 0x%1!X! (%1!d!)", m_stDaylightDate.wSecond->getData());
		daylight->addChild(
			m_stDaylightDate.wMilliseconds,
			L"wMilliseconds = 0x%1!X! (%1!d!)",
			m_stDaylightDate.wMilliseconds->getData());
	}
} // namespace smartview