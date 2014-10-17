#include "stdafx.h"
#include "..\stdafx.h"
#include "SmartViewParser.h"
#include "..\String.h"
#include "..\ParseProperty.h"

SmartViewParser::SmartViewParser(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	m_cbBin = cbBin;
	m_lpBin = lpBin;

	m_JunkDataSize = 0;
	m_JunkData = NULL;
}

SmartViewParser::~SmartViewParser() 
{
	delete[] m_JunkData;
}

_Check_return_ wstring SmartViewParser::JunkDataToString()
{
	if (!m_JunkDataSize || !m_JunkData) return L"";
	DebugPrint(DBGSmartView, _T("Had 0x%08X = %u bytes left over.\n"), (int)m_JunkDataSize, (UINT)m_JunkDataSize);
	wstring szJunkString = formatmessage(IDS_JUNKDATASIZE, m_JunkDataSize);
	SBinary sBin = { 0 };

	sBin.cb = (ULONG)m_JunkDataSize;
	sBin.lpb = m_JunkData;
	szJunkString += BinToHexString(&sBin, true);
	return szJunkString;
}

_Check_return_ wstring SmartViewParser::RTimeToString(DWORD rTime)
{
	wstring rTimeString;
	wstring rTimeAltString;
	FILETIME fTime = { 0 };
	LARGE_INTEGER liNumSec = { 0 };
	liNumSec.LowPart = rTime;
	// Resolution of RTime is in minutes, FILETIME is in 100 nanosecond intervals
	// Scale between the two is 10000000*60
	liNumSec.QuadPart = liNumSec.QuadPart * 10000000 * 60;
	fTime.dwLowDateTime = liNumSec.LowPart;
	fTime.dwHighDateTime = liNumSec.HighPart;
	FileTimeToString(&fTime, rTimeString, rTimeAltString);
	return rTimeString;
}