#include "stdafx.h"
#include "NickNameCache.h"
#include "String.h"
#include "PropertyStruct.h"

NickNameCache::NickNameCache()
{
	m_ulMajorVersion = 0;
	m_ulMinorVersion = 0;
	m_cRowCount = 0;
	m_lpRows = nullptr;
	m_cbEI = 0;
}

NickNameCache::~NickNameCache()
{
	if (m_cRowCount && m_lpRows)
	{
		for (DWORD i = 0; i < m_cRowCount; i++)
		{
			DeleteSPropVal(m_lpRows[i].cValues, m_lpRows[i].lpProps);
		}

		delete[] m_lpRows;
	}
}

void NickNameCache::Parse()
{
	m_Metadata1 = m_Parser.GetBYTES(4);
	m_Parser.GetDWORD(&m_ulMajorVersion);
	m_Parser.GetDWORD(&m_ulMinorVersion);
	m_Parser.GetDWORD(&m_cRowCount);

	if (m_cRowCount && m_cRowCount < _MaxEntriesEnormous)
		m_lpRows = new SRow[m_cRowCount];

	if (m_lpRows)
	{
		memset(m_lpRows, 0, sizeof(SRow)*m_cRowCount);

		for (DWORD i = 0; i < m_cRowCount; i++)
		{
			m_Parser.GetDWORD(&m_lpRows[i].cValues);

			if (m_lpRows[i].cValues && m_lpRows[i].cValues < _MaxEntriesSmall)
			{
				size_t cbBytesRead = 0;
				m_lpRows[i].lpProps = NickNameBinToSPropValue(
					m_lpRows[i].cValues);
				m_Parser.Advance(cbBytesRead);
			}
		}
	}

	m_Parser.GetDWORD(&m_cbEI);
	m_lpbEI = m_Parser.GetBYTES(m_cbEI, _MaxBytes);
	m_Metadata2 = m_Parser.GetBYTES(8);
}

// Caller allocates with new. Clean up with DeleteSPropVal.
_Check_return_ LPSPropValue NickNameCache::NickNameBinToSPropValue(DWORD dwPropCount)
{
	if (!dwPropCount || dwPropCount > _MaxEntriesSmall) return nullptr;

	auto pspvProperty = new SPropValue[dwPropCount];
	if (!pspvProperty) return nullptr;

	memset(pspvProperty, 0, sizeof(SPropValue)*dwPropCount);

	for (DWORD i = 0; i < dwPropCount; i++)
	{
		WORD PropType = 0;
		WORD PropID = 0;

		m_Parser.GetWORD(&PropType);
		m_Parser.GetWORD(&PropID);

		pspvProperty[i].ulPropTag = PROP_TAG(PropType, PropID);
		pspvProperty[i].dwAlignPad = 0;

		LARGE_INTEGER liTemp = { 0 };
		DWORD dwTemp = 0;
		m_Parser.GetDWORD(&dwTemp); // reserved
		m_Parser.GetLARGE_INTEGER(&liTemp); // union

		switch (PropType)
		{
		case PT_I2:
			pspvProperty[i].Value.i = static_cast<short int>(liTemp.LowPart);
			break;
		case PT_LONG:
			pspvProperty[i].Value.l = liTemp.LowPart;
			break;
		case PT_ERROR:
			pspvProperty[i].Value.err = liTemp.LowPart;
			break;
		case PT_R4:
			pspvProperty[i].Value.flt = static_cast<float>(liTemp.QuadPart);
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
			m_Parser.GetDWORD(&dwTemp);
			m_Parser.GetStringA(dwTemp, &pspvProperty[i].Value.lpszA);
			break;
		case PT_UNICODE:
			m_Parser.GetDWORD(&dwTemp);
			m_Parser.GetStringW(dwTemp / sizeof(WCHAR), &pspvProperty[i].Value.lpszW);
			break;
		case PT_CLSID:
			m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), reinterpret_cast<LPBYTE>(pspvProperty[i].Value.lpguid));
			break;
		case PT_BINARY:
			m_Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.bin.cb = dwTemp;
			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			m_Parser.GetBYTES(pspvProperty[i].Value.bin.cb, pspvProperty[i].Value.bin.cb, &pspvProperty[i].Value.bin.lpb);
			break;
		case PT_MV_BINARY:
			m_Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.MVbin.cValues = dwTemp;
			if (pspvProperty[i].Value.MVbin.cValues && pspvProperty[i].Value.MVbin.cValues < _MaxEntriesLarge)
			{
				pspvProperty[i].Value.MVbin.lpbin = new SBinary[dwTemp];
				if (pspvProperty[i].Value.MVbin.lpbin)
				{
					memset(pspvProperty[i].Value.MVbin.lpbin, 0, sizeof(SBinary)* dwTemp);
					for (ULONG j = 0; j < pspvProperty[i].Value.MVbin.cValues; j++)
					{
						m_Parser.GetDWORD(&dwTemp);
						pspvProperty[i].Value.MVbin.lpbin[j].cb = dwTemp;
						// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
						m_Parser.GetBYTES(pspvProperty[i].Value.MVbin.lpbin[j].cb,
							pspvProperty[i].Value.MVbin.lpbin[j].cb,
							&pspvProperty[i].Value.MVbin.lpbin[j].lpb);
					}
				}
			}
			break;
		case PT_MV_STRING8:
			m_Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.MVszA.cValues = dwTemp;
			if (pspvProperty[i].Value.MVszA.cValues && pspvProperty[i].Value.MVszA.cValues < _MaxEntriesLarge)
			{
				pspvProperty[i].Value.MVszA.lppszA = new CHAR*[dwTemp];
				if (pspvProperty[i].Value.MVszA.lppszA)
				{
					memset(pspvProperty[i].Value.MVszA.lppszA, 0, sizeof(CHAR*)* dwTemp);
					for (ULONG j = 0; j < pspvProperty[i].Value.MVszA.cValues; j++)
					{
						m_Parser.GetStringA(&pspvProperty[i].Value.MVszA.lppszA[j]);
					}
				}
			}
			break;
		case PT_MV_UNICODE:
			m_Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.MVszW.cValues = dwTemp;
			if (pspvProperty[i].Value.MVszW.cValues && pspvProperty[i].Value.MVszW.cValues < _MaxEntriesLarge)
			{
				pspvProperty[i].Value.MVszW.lppszW = new WCHAR*[dwTemp];
				if (pspvProperty[i].Value.MVszW.lppszW)
				{
					memset(pspvProperty[i].Value.MVszW.lppszW, 0, sizeof(WCHAR*)* dwTemp);
					for (ULONG j = 0; j < pspvProperty[i].Value.MVszW.cValues; j++)
					{
						m_Parser.GetStringW(&pspvProperty[i].Value.MVszW.lppszW[j]);
					}
				}
			}
			break;
		default:
			break;
		}
	}

	return pspvProperty;
}

_Check_return_ wstring NickNameCache::ToStringInternal()
{
	wstring szNickNameCache;

	szNickNameCache = formatmessage(IDS_NICKNAMEHEADER);
	szNickNameCache += BinToHexString(m_Metadata1, true);

	szNickNameCache += formatmessage(IDS_NICKNAMEROWCOUNT, m_ulMajorVersion, m_ulMinorVersion, m_cRowCount);

	if (m_cRowCount && m_lpRows)
	{
		for (DWORD i = 0; i < m_cRowCount; i++)
		{
			szNickNameCache += formatmessage(IDS_NICKNAMEROWS,
				i,
				m_lpRows[i].cValues);

			szNickNameCache += PropsToString(m_lpRows[i].cValues, m_lpRows[i].lpProps);
		}
	}

	szNickNameCache += formatmessage(IDS_NICKNAMEEXTRAINFO);
	szNickNameCache += BinToHexString(m_lpbEI, true);

	szNickNameCache += formatmessage(IDS_NICKNAMEFOOTER);
	szNickNameCache += BinToHexString(m_Metadata2, true);

	return szNickNameCache;
}