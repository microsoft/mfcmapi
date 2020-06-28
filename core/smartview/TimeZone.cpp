#include <core/stdafx.h>
#include <core/smartview/TimeZone.h>

namespace smartview
{
	void SYSTEMTIMEBlock::parse()
	{
		wYear = blockT<WORD>::parse(parser);
		wMonth = blockT<WORD>::parse(parser);
		wDayOfWeek = blockT<WORD>::parse(parser);
		wDay = blockT<WORD>::parse(parser);
		wHour = blockT<WORD>::parse(parser);
		wMinute = blockT<WORD>::parse(parser);
		wSecond = blockT<WORD>::parse(parser);
		wMilliseconds = blockT<WORD>::parse(parser);
	}

	void SYSTEMTIMEBlock::parseBlocks()
	{
		addChild(wYear, L"wYear = 0x%1!X! (%1!d!)", wYear->getData());
		addChild(wMonth, L"wMonth = 0x%1!X! (%1!d!)", wMonth->getData());
		addChild(wDayOfWeek, L"wDayOfWeek = 0x%1!X! (%1!d!)", wDayOfWeek->getData());
		addChild(wDay, L"wDay = 0x%1!X! (%1!d!)", wDay->getData());
		addChild(wHour, L"wHour = 0x%1!X! (%1!d!)", wHour->getData());
		addChild(wMinute, L"wMinute = 0x%1!X! (%1!d!)", wMinute->getData());
		addChild(wSecond, L"wSecond = 0x%1!X! (%1!d!)", wSecond->getData());
		addChild(wMilliseconds, L"wMilliseconds = 0x%1!X! (%1!d!)", wMilliseconds->getData());
	}

	void TimeZone::parse()
	{
		m_lBias = blockT<DWORD>::parse(parser);
		m_lStandardBias = blockT<DWORD>::parse(parser);
		m_lDaylightBias = blockT<DWORD>::parse(parser);
		m_wStandardYear = blockT<WORD>::parse(parser);
		m_stStandardDate = block::parse<SYSTEMTIMEBlock>(parser, false);
		m_wDaylightDate = blockT<WORD>::parse(parser);
		m_stDaylightDate = block::parse<SYSTEMTIMEBlock>(parser, false);
	}

	void TimeZone::parseBlocks()
	{
		setText(L"Time Zone");
		addChild(m_lBias, L"lBias = 0x%1!08X! (%1!d!)", m_lBias->getData());
		addChild(m_lStandardBias, L"lStandardBias = 0x%1!08X! (%1!d!)", m_lStandardBias->getData());
		addChild(m_lDaylightBias, L"lDaylightBias = 0x%1!08X! (%1!d!)", m_lDaylightBias->getData());
		addChild(m_wStandardYear, L"wStandardYear = 0x%1!04X! (%1!d!)", m_wStandardYear->getData());
		addChild(m_stStandardDate, L"stStandardDate");
		addChild(m_wDaylightDate, L"wDaylightDate = 0x%1!04X! (%1!d!)", m_wDaylightDate->getData());
		addChild(m_stDaylightDate, L"stDaylightDate");
	}
} // namespace smartview