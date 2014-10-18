#include "stdafx.h"
#include "..\stdafx.h"
#include "NickNameCache.h"
#include "..\String.h"
#include "..\ParseProperty.h"
#include "SmartView.h"

_Check_return_ LPSPropValue NickNameBinToSPropValue(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, DWORD dwPropCount, _Out_ size_t* lpcbBytesRead);

NickNameCache::NickNameCache(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	memset(m_Metadata1, 0, sizeof(m_Metadata1));
	m_ulMajorVersion = 0;
	m_ulMinorVersion = 0;
	m_cRowCount = 0;
	m_lpRows = 0;
	m_cbEI = 0;
	m_lpbEI = 0;
	memset(m_Metadata2, 0, sizeof(m_Metadata2));
}

NickNameCache::~NickNameCache()
{
	if (m_cRowCount && m_lpRows)
	{
		ULONG i = 0;

		for (i = 0; i < m_cRowCount; i++)
		{
			DeleteSPropVal(m_lpRows[i].cValues, m_lpRows[i].lpProps);
		}

		delete[] m_lpRows;
	}

	delete[] m_lpbEI;
}

void NickNameCache::Parse()
{
	if (!m_lpBin) return;

	m_Parser.GetBYTESNoAlloc(sizeof(m_Metadata1), sizeof(m_Metadata1), m_Metadata1);
	m_Parser.GetDWORD(&m_ulMajorVersion);
	m_Parser.GetDWORD(&m_ulMinorVersion);
	m_Parser.GetDWORD(&m_cRowCount);

	if (m_cRowCount && m_cRowCount < _MaxEntriesEnormous)
		m_lpRows = new SRow[m_cRowCount];

	if (m_lpRows)
	{
		memset(m_lpRows, 0, sizeof(SRow)*m_cRowCount);
		ULONG i = 0;

		for (i = 0; i < m_cRowCount; i++)
		{
			m_Parser.GetDWORD(&m_lpRows[i].cValues);

			if (m_lpRows[i].cValues && m_lpRows[i].cValues < _MaxEntriesSmall)
			{
				size_t cbBytesRead = 0;
				m_lpRows[i].lpProps = NickNameBinToSPropValue(
					(ULONG)m_Parser.RemainingBytes(),
					m_lpBin + m_Parser.GetCurrentOffset(),
					m_lpRows[i].cValues,
					&cbBytesRead);
				m_Parser.Advance(cbBytesRead);
			}
		}
	}

	m_Parser.GetDWORD(&m_cbEI);
	m_Parser.GetBYTES(m_cbEI, _MaxBytes, &m_lpbEI);
	m_Parser.GetBYTESNoAlloc(sizeof(m_Metadata2), sizeof(m_Metadata2), m_Metadata2);
}

_Check_return_ LPWSTR NickNameCache::ToString()
{
	Parse();

	wstring szNickNameCache;
	wstring szTmp;

	szNickNameCache = formatmessage(IDS_NICKNAMEHEADER);
	SBinary sBinMetadata = { 0 };
	sBinMetadata.cb = sizeof(m_Metadata1);
	sBinMetadata.lpb = m_Metadata1;
	szNickNameCache += BinToHexString(&sBinMetadata, true);

	szTmp = formatmessage(IDS_NICKNAMEROWCOUNT, m_ulMajorVersion, m_ulMinorVersion, m_cRowCount);
	szNickNameCache += szTmp;

	if (m_cRowCount && m_lpRows)
	{
		ULONG i = 0;
		for (i = 0; i < m_cRowCount; i++)
		{
			szTmp = formatmessage(IDS_NICKNAMEROWS,
				i,
				m_lpRows[i].cValues);
			szNickNameCache += szTmp;

			PropertyStruct psPropStruct = { 0 };
			psPropStruct.PropCount = m_lpRows[i].cValues;
			psPropStruct.Prop = m_lpRows[i].lpProps;

			LPWSTR szProps = PropertyStructToString(&psPropStruct);
			szNickNameCache += szProps;
			delete[] szProps;
		}
	}

	SBinary sBinEI = { 0 };
	szTmp = formatmessage(IDS_NICKNAMEEXTRAINFO);
	szNickNameCache += szTmp;

	sBinEI.cb = m_cbEI;
	sBinEI.lpb = m_lpbEI;
	szNickNameCache += BinToHexString(&sBinEI, true);

	szTmp = formatmessage(IDS_NICKNAMEFOOTER);
	szNickNameCache += szTmp;

	sBinMetadata.cb = sizeof(m_Metadata2);
	sBinMetadata.lpb = m_Metadata2;
	szNickNameCache += BinToHexString(&sBinMetadata, true);

	szNickNameCache += JunkDataToString();

	return wstringToLPWSTR(szNickNameCache);
}

// Caller allocates with new. Clean up with DeleteSPropVal.
_Check_return_ LPSPropValue NickNameBinToSPropValue(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, DWORD dwPropCount, _Out_ size_t* lpcbBytesRead)
{
	if (!lpBin) return NULL;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	if (!dwPropCount || dwPropCount > _MaxEntriesSmall) return NULL;

	LPSPropValue pspvProperty = new SPropValue[dwPropCount];
	if (!pspvProperty) return NULL;

	memset(pspvProperty, 0, sizeof(SPropValue)*dwPropCount);
	CBinaryParser Parser(cbBin, lpBin);

	DWORD i = 0;

	for (i = 0; i < dwPropCount; i++)
	{
		WORD PropType = 0;
		WORD PropID = 0;

		Parser.GetWORD(&PropType);
		Parser.GetWORD(&PropID);

		pspvProperty[i].ulPropTag = PROP_TAG(PropType, PropID);
		pspvProperty[i].dwAlignPad = 0;

		LARGE_INTEGER liTemp = { 0 };
		DWORD dwTemp = 0;
		Parser.GetDWORD(&dwTemp); // reserved
		Parser.GetLARGE_INTEGER(&liTemp); // union

		switch (PropType)
		{
		case PT_I2:
			pspvProperty[i].Value.i = (short int)liTemp.LowPart;
			break;
		case PT_LONG:
			pspvProperty[i].Value.l = liTemp.LowPart;
			break;
		case PT_ERROR:
			pspvProperty[i].Value.err = liTemp.LowPart;
			break;
		case PT_R4:
			pspvProperty[i].Value.flt = (float)liTemp.QuadPart;
			break;
		case PT_DOUBLE:
			pspvProperty[i].Value.dbl = liTemp.LowPart;
			break;
		case PT_BOOLEAN:
			pspvProperty[i].Value.b = liTemp.LowPart ? true : false;
			break;
		case PT_SYSTIME:
			pspvProperty[i].Value.ft.dwHighDateTime = liTemp.HighPart;
			pspvProperty[i].Value.ft.dwLowDateTime = liTemp.LowPart;
			break;
		case PT_I8:
			pspvProperty[i].Value.li = liTemp;
			break;
		case PT_STRING8:
			Parser.GetDWORD(&dwTemp);
			Parser.GetStringA(dwTemp, &pspvProperty[i].Value.lpszA);
			break;
		case PT_UNICODE:
			Parser.GetDWORD(&dwTemp);
			Parser.GetStringW(dwTemp / sizeof(WCHAR), &pspvProperty[i].Value.lpszW);
			break;
		case PT_CLSID:
			Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)pspvProperty[i].Value.lpguid);
			break;
		case PT_BINARY:
			Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.bin.cb = dwTemp;
			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			Parser.GetBYTES(pspvProperty[i].Value.bin.cb, pspvProperty[i].Value.bin.cb, &pspvProperty[i].Value.bin.lpb);
			break;
		case PT_MV_BINARY:
			Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.MVbin.cValues = dwTemp;
			if (pspvProperty[i].Value.MVbin.cValues && pspvProperty[i].Value.MVbin.cValues < _MaxEntriesLarge)
			{
				pspvProperty[i].Value.MVbin.lpbin = new SBinary[dwTemp];
				if (pspvProperty[i].Value.MVbin.lpbin)
				{
					memset(pspvProperty[i].Value.MVbin.lpbin, 0, sizeof(SBinary)* dwTemp);
					DWORD j = 0;
					for (j = 0; j < pspvProperty[i].Value.MVbin.cValues; j++)
					{
						Parser.GetDWORD(&dwTemp);
						pspvProperty[i].Value.MVbin.lpbin[j].cb = dwTemp;
						// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
						Parser.GetBYTES(pspvProperty[i].Value.MVbin.lpbin[j].cb,
							pspvProperty[i].Value.MVbin.lpbin[j].cb,
							&pspvProperty[i].Value.MVbin.lpbin[j].lpb);
					}
				}
			}
			break;
		case PT_MV_STRING8:
			Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.MVszA.cValues = dwTemp;
			if (pspvProperty[i].Value.MVszA.cValues && pspvProperty[i].Value.MVszA.cValues < _MaxEntriesLarge)
			{
				pspvProperty[i].Value.MVszA.lppszA = new CHAR*[dwTemp];
				if (pspvProperty[i].Value.MVszA.lppszA)
				{
					memset(pspvProperty[i].Value.MVszA.lppszA, 0, sizeof(CHAR*)* dwTemp);
					DWORD j = 0;
					for (j = 0; j < pspvProperty[i].Value.MVszA.cValues; j++)
					{
						Parser.GetStringA(&pspvProperty[i].Value.MVszA.lppszA[j]);
					}
				}
			}
			break;
		case PT_MV_UNICODE:
			Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.MVszW.cValues = dwTemp;
			if (pspvProperty[i].Value.MVszW.cValues && pspvProperty[i].Value.MVszW.cValues < _MaxEntriesLarge)
			{
				pspvProperty[i].Value.MVszW.lppszW = new WCHAR*[dwTemp];
				if (pspvProperty[i].Value.MVszW.lppszW)
				{
					memset(pspvProperty[i].Value.MVszW.lppszW, 0, sizeof(WCHAR*)* dwTemp);
					DWORD j = 0;
					for (j = 0; j < pspvProperty[i].Value.MVszW.cValues; j++)
					{
						Parser.GetStringW(&pspvProperty[i].Value.MVszW.lppszW[j]);
					}
				}
			}
			break;
		default:
			break;
		}
	}

	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();
	return pspvProperty;
}