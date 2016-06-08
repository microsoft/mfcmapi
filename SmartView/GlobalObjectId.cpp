#include "stdafx.h"
#include "..\stdafx.h"
#include "GlobalObjectId.h"
#include "..\String.h"
#include "..\InterpretProp.h"
#include "..\InterpretProp2.h"
#include "..\ExtraPropTags.h"

GlobalObjectId::GlobalObjectId(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	memset(m_Id, 0, sizeof(m_Id));
	m_Year = 0;
	m_Month = 0;
	m_Day = 0;
	m_CreationTime = { 0 };
	m_X = { 0 };
	m_dwSize = 0;
	m_lpData = NULL;
}

GlobalObjectId::~GlobalObjectId()
{
	delete[] m_lpData;
}

void GlobalObjectId::Parse()
{
	m_Parser.GetBYTESNoAlloc(sizeof(m_Id), sizeof(m_Id), (LPBYTE)&m_Id);
	BYTE b1 = NULL;
	BYTE b2 = NULL;
	m_Parser.GetBYTE(&b1);
	m_Parser.GetBYTE(&b2);
	m_Year = (WORD)((b1 << 8) | b2);
	m_Parser.GetBYTE(&m_Month);
	m_Parser.GetBYTE(&m_Day);
	m_Parser.GetLARGE_INTEGER((LARGE_INTEGER*)&m_CreationTime);
	m_Parser.GetLARGE_INTEGER(&m_X);
	m_Parser.GetDWORD(&m_dwSize);
	m_Parser.GetBYTES(m_dwSize, _MaxBytes, &m_lpData);
}

_Check_return_ wstring GlobalObjectId::ToStringInternal()
{
	wstring szGlobalObjectId;

	szGlobalObjectId = formatmessage(IDS_GLOBALOBJECTIDHEADER);

	SBinary sBin = { 0 };
	sBin.cb = sizeof(m_Id);
	sBin.lpb = m_Id;
	szGlobalObjectId += BinToHexString(&sBin, true);

	wstring szFlags = InterpretFlags(flagGlobalObjectIdMonth, m_Month);

	wstring PropString;
	wstring AltPropString;
	FileTimeToString(&m_CreationTime, PropString, AltPropString);
	szGlobalObjectId += formatmessage(IDS_GLOBALOBJECTIDDATA1,
		m_Year,
		m_Month, szFlags.c_str(),
		m_Day,
		m_CreationTime.dwHighDateTime, m_CreationTime.dwLowDateTime, PropString.c_str(),
		m_X.HighPart, m_X.LowPart,
		m_dwSize);

	if (m_dwSize && m_lpData)
	{
		szGlobalObjectId += formatmessage(IDS_GLOBALOBJECTIDDATA2);
		sBin.cb = m_dwSize;
		sBin.lpb = m_lpData;
		szGlobalObjectId += BinToHexString(&sBin, true);
	}

	return szGlobalObjectId;
}