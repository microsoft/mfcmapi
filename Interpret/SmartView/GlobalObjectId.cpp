#include "stdafx.h"
#include "GlobalObjectId.h"
#include <Interpret/InterpretProp.h>
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>

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
	auto b1 = m_Parser.Get<BYTE>();
	auto b2 = m_Parser.Get<BYTE>();
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

_Check_return_ wstring GlobalObjectId::ToStringInternal()
{
	wstring szGlobalObjectId;

	szGlobalObjectId = formatmessage(IDS_GLOBALOBJECTIDHEADER);

	szGlobalObjectId += BinToHexString(m_Id, true);
	szGlobalObjectId += L" = ";
	if (equal(m_Id.begin(), m_Id.end(), s_rgbSPlus))
	{
		szGlobalObjectId += formatmessage(IDS_GLOBALOBJECTSPLUS);
	}
	else
	{
		szGlobalObjectId += formatmessage(IDS_UNKNOWNGUID);
	}

	auto szFlags = InterpretFlags(flagGlobalObjectIdMonth, m_Month);

	wstring PropString;
	wstring AltPropString;
	FileTimeToString(m_CreationTime, PropString, AltPropString);
	szGlobalObjectId += formatmessage(IDS_GLOBALOBJECTIDDATA1,
		m_Year,
		m_Month, szFlags.c_str(),
		m_Day,
		m_CreationTime.dwHighDateTime, m_CreationTime.dwLowDateTime, PropString.c_str(),
		m_X.HighPart, m_X.LowPart,
		m_dwSize);

	if (m_lpData.size())
	{
		szGlobalObjectId += formatmessage(IDS_GLOBALOBJECTIDDATA2);
		szGlobalObjectId += BinToHexString(m_lpData, true);
	}

	return szGlobalObjectId;
}