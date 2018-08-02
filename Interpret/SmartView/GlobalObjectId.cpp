#include <StdAfx.h>
#include <Interpret/SmartView/GlobalObjectId.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	// clang-format off
	static const BYTE s_rgbSPlus[] =
	{
		0x04, 0x00, 0x00, 0x00,
		0x82, 0x00, 0xE0, 0x00,
		0x74, 0xC5, 0xB7, 0x10,
		0x1A, 0x82, 0xE0, 0x08,
	};
	// clang-format on

	void GlobalObjectId::Parse()
	{
		addHeader(L"Global Object ID:\r\n");
		addHeader(L"Byte Array ID = ");

		auto id = m_Parser.GetBYTES(16);
		addBytes(id);

		if (equal(id.begin(), id.end(), s_rgbSPlus))
		{
			addHeader(L" = s_rgbSPlus\r\n");
		}
		else
		{
			addHeader(L" = Unknown GUID\r\n");
		}

		const auto b1 = m_Parser.Get<BYTE>();
		const auto b2 = m_Parser.Get<BYTE>();
		const auto year = static_cast<WORD>(b1 << 8 | b2);
		addData(2 * sizeof BYTE, strings::formatmessage(L"Year: 0x%1!04X! = %1!d!\r\n", year));

		const auto month = m_Parser.Get<BYTE>();
		const auto szFlags = interpretprop::InterpretFlags(flagGlobalObjectIdMonth, month);
		addData(sizeof BYTE, strings::formatmessage(L"Month: 0x%1!02X! = %1!d! = %2!ws!\r\n", month, szFlags.c_str()));

		const auto day = m_Parser.Get<BYTE>();
		addData(sizeof BYTE, strings::formatmessage(L"Day: 0x%1!02X! = %1!d!\r\n", day));

		const auto creationTime = m_Parser.Get<FILETIME>();
		std::wstring propString;
		std::wstring altPropString;
		strings::FileTimeToString(creationTime, propString, altPropString);
		addData(
			sizeof FILETIME,
			strings::formatmessage(
				L"Creation Time = 0x%1!08X!:0x%2!08X! = %3!ws!\r\n",
				creationTime.dwHighDateTime,
				creationTime.dwLowDateTime,
				propString.c_str()));

		const auto x = m_Parser.Get<LARGE_INTEGER>();
		addData(sizeof LARGE_INTEGER, strings::formatmessage(L"X: 0x%1!08X!:0x%2!08X!\r\n", x.HighPart, x.LowPart));

		const auto dwSize = m_Parser.Get<DWORD>();
		addData(sizeof DWORD, strings::formatmessage(L"Size: 0x%1!02X! = %1!d!\r\n", dwSize));

		const auto lpData = m_Parser.GetBYTES(dwSize, _MaxBytes);
		if (lpData.size())
		{
			addHeader(L"Data = ");
			addBytes(lpData);
		}
	}
}