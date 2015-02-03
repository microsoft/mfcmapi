#include "stdafx.h"
#include "..\stdafx.h"
#include "SmartViewParser.h"
#include "..\String.h"
#include "..\ParseProperty.h"

SmartViewParser::SmartViewParser(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	m_bParsed = false;
	m_bEnableJunk = true;
	m_cbBin = cbBin;
	m_lpBin = lpBin;

	m_Parser.Init(m_cbBin, m_lpBin);
}

SmartViewParser::~SmartViewParser()
{
}

void SmartViewParser::DisableJunkParsing()
{
	m_bEnableJunk = false;
}

size_t SmartViewParser::GetCurrentOffset()
{
	return m_Parser.GetCurrentOffset();
}

void SmartViewParser::EnsureParsed()
{
	if (m_bParsed || m_Parser.Empty()) return;
	Parse();
	m_bParsed = true;
}

_Check_return_ wstring SmartViewParser::ToString()
{
	if (m_Parser.Empty()) return L"";
	EnsureParsed();

	wstring szParsedString = ToStringInternal();

	if (m_bEnableJunk)
	{
		szParsedString += JunkDataToString();
	}

	return szParsedString;
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

// TODO: Eliminate this
_Check_return_ wstring SmartViewParser::JunkDataToString(size_t cbJunkData, _In_count_(cbJunkData) LPBYTE lpJunkData)
{
	if (!cbJunkData || !lpJunkData) return L"";
	DebugPrint(DBGSmartView, _T("Had 0x%08X = %u bytes left over.\n"), (int)cbJunkData, (UINT)cbJunkData);
	wstring szTmp;
	SBinary sBin = { 0 };

	sBin.cb = (ULONG)cbJunkData;
	sBin.lpb = lpJunkData;
	szTmp = formatmessage(IDS_JUNKDATASIZE,
		cbJunkData);
	szTmp += BinToHexString(&sBin, true).c_str();
	return szTmp;
}

// Caller allocates with new. Clean up with DeleteSPropVal.
_Check_return_ LPSPropValue SmartViewParser::BinToSPropValue(DWORD dwPropCount, bool bStringPropsExcludeLength)
{
	if (!dwPropCount || dwPropCount > _MaxEntriesSmall) return NULL;
	LPSPropValue pspvProperty = new SPropValue[dwPropCount];
	if (!pspvProperty) return NULL;
	memset(pspvProperty, 0, sizeof(SPropValue)*dwPropCount);

	DWORD i = 0;

	for (i = 0; i < dwPropCount; i++)
	{
		WORD PropType = 0;
		WORD PropID = 0;

		m_Parser.GetWORD(&PropType);
		m_Parser.GetWORD(&PropID);

		pspvProperty[i].ulPropTag = PROP_TAG(PropType, PropID);
		pspvProperty[i].dwAlignPad = 0;

		DWORD dwTemp = 0;
		WORD wTemp = 0;
		switch (PropType)
		{
		case PT_LONG:
			m_Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.l = dwTemp;
			break;
		case PT_ERROR:
			m_Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.err = dwTemp;
			break;
		case PT_BOOLEAN:
			m_Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.b = wTemp;
			break;
		case PT_UNICODE:
			if (bStringPropsExcludeLength)
			{
				m_Parser.GetStringW(&pspvProperty[i].Value.lpszW);
			}
			else
			{
				// This is apparently a cb...
				m_Parser.GetWORD(&wTemp);
				m_Parser.GetStringW(wTemp / sizeof(WCHAR), &pspvProperty[i].Value.lpszW);
			}
			break;
		case PT_STRING8:
			if (bStringPropsExcludeLength)
			{
				m_Parser.GetStringA(&pspvProperty[i].Value.lpszA);
			}
			else
			{
				// This is apparently a cb...
				m_Parser.GetWORD(&wTemp);
				m_Parser.GetStringA(wTemp, &pspvProperty[i].Value.lpszA);
			}
			break;
		case PT_SYSTIME:
			m_Parser.GetDWORD(&pspvProperty[i].Value.ft.dwHighDateTime);
			m_Parser.GetDWORD(&pspvProperty[i].Value.ft.dwLowDateTime);
			break;
		case PT_BINARY:
			m_Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.bin.cb = wTemp;
			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			m_Parser.GetBYTES(pspvProperty[i].Value.bin.cb, pspvProperty[i].Value.bin.cb, &pspvProperty[i].Value.bin.lpb);
			break;
		case PT_MV_STRING8:
			m_Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.MVszA.cValues = wTemp;
			pspvProperty[i].Value.MVszA.lppszA = new CHAR*[wTemp];
			if (pspvProperty[i].Value.MVszA.lppszA)
			{
				memset(pspvProperty[i].Value.MVszA.lppszA, 0, sizeof(CHAR*)* wTemp);
				DWORD j = 0;
				for (j = 0; j < pspvProperty[i].Value.MVszA.cValues; j++)
				{
					m_Parser.GetStringA(&pspvProperty[i].Value.MVszA.lppszA[j]);
				}
			}
			break;
		case PT_MV_UNICODE:
			m_Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.MVszW.cValues = wTemp;
			pspvProperty[i].Value.MVszW.lppszW = new WCHAR*[wTemp];
			if (pspvProperty[i].Value.MVszW.lppszW)
			{
				memset(pspvProperty[i].Value.MVszW.lppszW, 0, sizeof(WCHAR*)* wTemp);
				DWORD j = 0;
				for (j = 0; j < pspvProperty[i].Value.MVszW.cValues; j++)
				{
					m_Parser.GetStringW(&pspvProperty[i].Value.MVszW.lppszW[j]);
				}
			}
			break;
		default:
			break;
		}
	}

	return pspvProperty;
}