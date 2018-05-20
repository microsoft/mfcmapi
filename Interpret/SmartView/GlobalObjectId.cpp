#include "StdAfx.h"
#include "GlobalObjectId.h"
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	GlobalObjectId::GlobalObjectId()
	{
		m_Year = 0;
		m_Month = 0;
		m_Day = 0;
		m_CreationTime = { 0 };
		m_X = { 0 };
		m_dwSize = 0;
	}

	void GlobalObjectId::Parse()
	{
		m_Id = m_Parser.GetBYTES(16);
		const auto b1 = m_Parser.Get<BYTE>();
		const auto b2 = m_Parser.Get<BYTE>();
		m_Year = static_cast<WORD>(b1 << 8 | b2);
		m_Month = m_Parser.Get<BYTE>();
		m_Day = m_Parser.Get<BYTE>();
		m_CreationTime = m_Parser.Get<FILETIME>();
		m_X = m_Parser.Get<LARGE_INTEGER>();
		m_dwSize = m_Parser.Get<DWORD>();
		m_lpData = m_Parser.GetBYTES(m_dwSize, _MaxBytes);
	}

	static const BYTE s_rgbSPlus[] =
	{
		0x04, 0x00, 0x00, 0x00,
		0x82, 0x00, 0xE0, 0x00,
		0x74, 0xC5, 0xB7, 0x10,
		0x1A, 0x82, 0xE0, 0x08,
	};

	_Check_return_ std::wstring GlobalObjectId::ToStringInternal()
	{
		auto szGlobalObjectId = strings::formatmessage(IDS_GLOBALOBJECTIDHEADER);

		szGlobalObjectId += strings::BinToHexString(m_Id, true);
		szGlobalObjectId += L" = ";
		if (equal(m_Id.begin(), m_Id.end(), s_rgbSPlus))
		{
			szGlobalObjectId += strings::formatmessage(IDS_GLOBALOBJECTSPLUS);
		}
		else
		{
			szGlobalObjectId += strings::formatmessage(IDS_UNKNOWNGUID);
		}

		auto szFlags = interpretprop::InterpretFlags(flagGlobalObjectIdMonth, m_Month);

		std::wstring PropString;
		std::wstring AltPropString;
		strings::FileTimeToString(m_CreationTime, PropString, AltPropString);
		szGlobalObjectId += strings::formatmessage(IDS_GLOBALOBJECTIDDATA1,
			m_Year,
			m_Month, szFlags.c_str(),
			m_Day,
			m_CreationTime.dwHighDateTime, m_CreationTime.dwLowDateTime, PropString.c_str(),
			m_X.HighPart, m_X.LowPart,
			m_dwSize);

		if (m_lpData.size())
		{
			szGlobalObjectId += strings::formatmessage(IDS_GLOBALOBJECTIDDATA2);
			szGlobalObjectId += strings::BinToHexString(m_lpData, true);
		}

		return szGlobalObjectId;
	}
}