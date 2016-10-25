#include "stdafx.h"
#include "SmartViewParser.h"
#include "String.h"
#include <algorithm>

SmartViewParser::SmartViewParser()
{
	m_bParsed = false;
	m_bEnableJunk = true;
	m_cbBin = 0;
	m_lpBin = nullptr;
}

void SmartViewParser::Init(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	m_cbBin = cbBin;
	m_lpBin = lpBin;
	m_Parser.Init(m_cbBin, m_lpBin);
}

void SmartViewParser::DisableJunkParsing()
{
	m_bEnableJunk = false;
}

size_t SmartViewParser::GetCurrentOffset() const
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

	auto szParsedString = ToStringInternal();

	if (m_bEnableJunk)
	{
		szParsedString += JunkDataToString(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
	}

	// If we built a string with emedded nulls in it, replace them with dots.
	std::replace_if(szParsedString.begin(), szParsedString.end(), [](const WCHAR & chr)
	{
		return chr == L'\0';
	}, L'.');

	return szParsedString;
}

_Check_return_ wstring SmartViewParser::JunkDataToString(const vector<BYTE>& lpJunkData) const
{
	if (!lpJunkData.size()) return emptystring;
	DebugPrint(DBGSmartView, L"Had 0x%08X = %u bytes left over.\n", static_cast<int>(lpJunkData.size()), static_cast<UINT>(lpJunkData.size()));
	auto szJunk = formatmessage(IDS_JUNKDATASIZE, lpJunkData.size());
	szJunk += BinToHexString(lpJunkData, true);
	return szJunk;
}

_Check_return_ wstring SmartViewParser::JunkDataToString(size_t cbJunkData, _In_count_(cbJunkData) LPBYTE lpJunkData) const
{
	if (!cbJunkData || !lpJunkData) return L"";
	DebugPrint(DBGSmartView, L"Had 0x%08X = %u bytes left over.\n", static_cast<int>(cbJunkData), static_cast<UINT>(cbJunkData));
	SBinary sBin = { 0 };
	sBin.cb = static_cast<ULONG>(cbJunkData);
	sBin.lpb = lpJunkData;
	auto szJunk = formatmessage(IDS_JUNKDATASIZE, cbJunkData);
	szJunk += BinToHexString(&sBin, true);
	return szJunk;
}

LPBYTE SmartViewParser::GetBYTES(size_t cbBytes, size_t cbMaxBytes)
{
	m_binCache.push_back(m_Parser.GetBYTES(cbBytes, cbMaxBytes));
	return m_binCache.back().data();
}

LPSTR SmartViewParser::GetStringA(size_t cchChar)
{
	m_stringCache.push_back(m_Parser.GetStringA(cchChar));
	return const_cast<LPSTR>(m_stringCache.back().data());
}

LPWSTR SmartViewParser::GetStringW(size_t cchChar)
{
	m_wstringCache.push_back(m_Parser.GetStringW(cchChar));
	return const_cast<LPWSTR>(m_wstringCache.back().data());
}

LPBYTE SmartViewParser::Allocate(size_t cbBytes)
{
	vector<BYTE> bin;
	bin.resize(cbBytes);
	m_binCache.push_back(bin);
	return m_binCache.back().data();
}

LPBYTE SmartViewParser::AllocateArray(size_t cArray, size_t cbEntry)
{
	return Allocate(cArray * cbEntry);
}


_Check_return_ LPSPropValue SmartViewParser::BinToSPropValue(DWORD dwPropCount, bool bStringPropsExcludeLength)
{
	if (!dwPropCount || dwPropCount > _MaxEntriesSmall) return nullptr;
	auto pspvProperty = reinterpret_cast<SPropValue*>(AllocateArray(dwPropCount, sizeof SPropValue));
	if (!pspvProperty) return nullptr;
	auto ulCurrOffset = m_Parser.GetCurrentOffset();

	for (DWORD i = 0; i < dwPropCount; i++)
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
				pspvProperty[i].Value.lpszW = GetStringW();
			}
			else
			{
				// This is apparently a cb...
				m_Parser.GetWORD(&wTemp);
				pspvProperty[i].Value.lpszW = GetStringW(wTemp / sizeof(WCHAR));
			}
			break;
		case PT_STRING8:
			if (bStringPropsExcludeLength)
			{
				pspvProperty[i].Value.lpszA = GetStringA();
			}
			else
			{
				// This is apparently a cb...
				m_Parser.GetWORD(&wTemp);
				pspvProperty[i].Value.lpszA = GetStringA(wTemp);
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
			pspvProperty[i].Value.bin.lpb = GetBYTES(pspvProperty[i].Value.bin.cb);
			break;
		case PT_MV_STRING8:
			m_Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.MVszA.cValues = wTemp;
			pspvProperty[i].Value.MVszA.lppszA = reinterpret_cast<LPSTR*>(AllocateArray(wTemp, sizeof LPVOID));
			if (pspvProperty[i].Value.MVszA.lppszA)
			{
				for (ULONG j = 0; j < pspvProperty[i].Value.MVszA.cValues; j++)
				{
					pspvProperty[i].Value.MVszA.lppszA[j] = GetStringA();
				}
			}
			break;
		case PT_MV_UNICODE:
			m_Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.MVszW.cValues = wTemp;
			pspvProperty[i].Value.MVszW.lppszW = reinterpret_cast<LPWSTR*>(AllocateArray(wTemp, sizeof LPVOID));
			if (pspvProperty[i].Value.MVszW.lppszW)
			{
				for (ULONG j = 0; j < pspvProperty[i].Value.MVszW.cValues; j++)
				{
					pspvProperty[i].Value.MVszW.lppszW[j] = GetStringW();
				}
			}
			break;
		default:
			break;
		}
	}

	if (ulCurrOffset == m_Parser.GetCurrentOffset())
	{
		// We didn't advance at all - we should return nothing.
		delete[] pspvProperty;
		pspvProperty = nullptr;
	}

	return pspvProperty;
}