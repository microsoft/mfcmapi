#include <core/stdafx.h>
#include <core/smartview/TimeZone.h>

namespace smartview
{
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

	void TimeZone::ParseBlocks()
	{
		setRoot(L"Time Zone: \r\n");
		addBlock(m_lBias, L"lBias = 0x%1!08X! (%1!d!)\r\n", m_lBias.getData());
		addBlock(m_lStandardBias, L"lStandardBias = 0x%1!08X! (%1!d!)\r\n", m_lStandardBias.getData());
		addBlock(m_lDaylightBias, L"lDaylightBias = 0x%1!08X! (%1!d!)\r\n", m_lDaylightBias.getData());
		addBlankLine();
		addBlock(m_wStandardYear, L"wStandardYear = 0x%1!04X! (%1!d!)\r\n", m_wStandardYear.getData());
		addBlock(
			m_stStandardDate.wYear, L"stStandardDate.wYear = 0x%1!X! (%1!d!)\r\n", m_stStandardDate.wYear.getData());
		addBlock(
			m_stStandardDate.wMonth, L"stStandardDate.wMonth = 0x%1!X! (%1!d!)\r\n", m_stStandardDate.wMonth.getData());
		addBlock(
			m_stStandardDate.wDayOfWeek,
			L"stStandardDate.wDayOfWeek = 0x%1!X! (%1!d!)\r\n",
			m_stStandardDate.wDayOfWeek.getData());
		addBlock(m_stStandardDate.wDay, L"stStandardDate.wDay = 0x%1!X! (%1!d!)\r\n", m_stStandardDate.wDay.getData());
		addBlock(
			m_stStandardDate.wHour, L"stStandardDate.wHour = 0x%1!X! (%1!d!)\r\n", m_stStandardDate.wHour.getData());
		addBlock(
			m_stStandardDate.wMinute,
			L"stStandardDate.wMinute = 0x%1!X! (%1!d!)\r\n",
			m_stStandardDate.wMinute.getData());
		addBlock(
			m_stStandardDate.wSecond,
			L"stStandardDate.wSecond = 0x%1!X! (%1!d!)\r\n",
			m_stStandardDate.wSecond.getData());
		addBlock(
			m_stStandardDate.wMilliseconds,
			L"stStandardDate.wMilliseconds = 0x%1!X! (%1!d!)\r\n",
			m_stStandardDate.wMilliseconds.getData());
		addBlankLine();
		addBlock(m_wDaylightDate, L"wDaylightDate = 0x%1!04X! (%1!d!)\r\n", m_wDaylightDate.getData());
		addBlock(
			m_stDaylightDate.wYear, L"stDaylightDate.wYear = 0x%1!X! (%1!d!)\r\n", m_stDaylightDate.wYear.getData());
		addBlock(
			m_stDaylightDate.wMonth, L"stDaylightDate.wMonth = 0x%1!X! (%1!d!)\r\n", m_stDaylightDate.wMonth.getData());
		addBlock(
			m_stDaylightDate.wDayOfWeek,
			L"stDaylightDate.wDayOfWeek = 0x%1!X! (%1!d!)\r\n",
			m_stDaylightDate.wDayOfWeek.getData());
		addBlock(m_stDaylightDate.wDay, L"stDaylightDate.wDay = 0x%1!X! (%1!d!)\r\n", m_stDaylightDate.wDay.getData());
		addBlock(
			m_stDaylightDate.wHour, L"stDaylightDate.wHour = 0x%1!X! (%1!d!)\r\n", m_stDaylightDate.wHour.getData());
		addBlock(
			m_stDaylightDate.wMinute,
			L"stDaylightDate.wMinute = 0x%1!X! (%1!d!)\r\n",
			m_stDaylightDate.wMinute.getData());
		addBlock(
			m_stDaylightDate.wSecond,
			L"stDaylightDate.wSecond = 0x%1!X! (%1!d!)\r\n",
			m_stDaylightDate.wSecond.getData());
		addBlock(
			m_stDaylightDate.wMilliseconds,
			L"stDaylightDate.wMilliseconds = 0x%1!X! (%1!d!)",
			m_stDaylightDate.wMilliseconds.getData());
	}
} // namespace smartview