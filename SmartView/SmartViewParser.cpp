#include "stdafx.h"
#include "..\stdafx.h"
#include "SmartViewParser.h"
#include "..\String.h"
#include "..\ParseProperty.h"

SmartViewParser::SmartViewParser(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	m_cbBin = cbBin;
	m_lpBin = lpBin;

	m_Parser.Init(m_cbBin, m_lpBin);
}

SmartViewParser::~SmartViewParser()
{
}

_Check_return_ wstring SmartViewParser::JunkDataToString()
{
	LPBYTE junkData = NULL;
	size_t junkDataSize = m_Parser.GetRemainingData(&junkData);
	wstring szJunkString;

	if (junkDataSize && junkData)
	{
		DebugPrint(DBGSmartView, _T("Had 0x%08X = %u bytes left over.\n"), (int)junkDataSize, (UINT)junkDataSize);
		szJunkString = formatmessage(IDS_JUNKDATASIZE, junkDataSize);
		SBinary sBin = { 0 };

		sBin.cb = (ULONG)junkDataSize;
		sBin.lpb = junkData;
		szJunkString += BinToHexString(&sBin, true);
		
	}

	if (junkData) delete[] junkData;

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