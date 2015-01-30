#include "stdafx.h"
#include "..\stdafx.h"
#include "PropertyStruct.h"
#include "..\String.h"
#include "..\InterpretProp.h"
#include "..\InterpretProp2.h"

PropertyStruct::PropertyStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_PropCount = 0;
	m_Prop = NULL;
}

PropertyStruct::~PropertyStruct()
{
	DeleteSPropVal(m_PropCount, m_Prop);
}

void PropertyStruct::Parse()
{
	// Have to count how many properties are here.
	// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
	size_t stBookmark = m_Parser.GetCurrentOffset();

	DWORD dwPropCount = 0;

	for (;;)
	{
		LPSPropValue lpProp = BinToSPropValue(1, false);
		if (lpProp)
		{
			dwPropCount++;
			DeleteSPropVal(1, lpProp);
		}
		else
		{
			break;
		}

		if (!m_Parser.RemainingBytes()) break;
	}

	m_Parser.SetCurrentOffset(stBookmark); // We're done with our first pass, restore the bookmark

	m_PropCount = dwPropCount;
	m_Prop = BinToSPropValue(dwPropCount, false);
}

_Check_return_ wstring PropertyStruct::ToStringInternal()
{
	return PropsToString(m_PropCount, m_Prop);
}

_Check_return_ wstring PropsToString(DWORD PropCount, LPSPropValue Prop)
{
	wstring szProperty;
	DWORD i = 0;

	if (Prop)
	{
		for (i = 0; i < PropCount; i++)
		{
			wstring PropString;
			wstring AltPropString;

			szProperty += formatmessage(IDS_PROPERTYDATAHEADER,
				i,
				Prop[i].ulPropTag);

			LPTSTR szExactMatches = NULL;
			LPTSTR szPartialMatches = NULL;
			PropTagToPropName(Prop[i].ulPropTag, false, &szExactMatches, &szPartialMatches);
			if (!IsNullOrEmpty(szExactMatches))
			{
				szProperty += formatmessage(IDS_PROPERTYDATAEXACTMATCHES,
					LPTSTRToWstring(szExactMatches).c_str());
			}

			if (!IsNullOrEmpty(szPartialMatches))
			{
				szProperty += formatmessage(IDS_PROPERTYDATAPARTIALMATCHES,
					LPTSTRToWstring(szPartialMatches).c_str());
			}

			delete[] szExactMatches;
			delete[] szPartialMatches;

			InterpretProp(&Prop[i], &PropString, &AltPropString);
			szProperty += formatmessage(IDS_PROPERTYDATA,
				PropString.c_str(),
				AltPropString.c_str());

			wstring szSmartView = InterpretPropSmartView(
				&Prop[i],
				NULL,
				NULL,
				NULL,
				false,
				false);

			if (!szSmartView.empty())
			{
				szProperty += formatmessage(IDS_PROPERTYDATASMARTVIEW,
					szSmartView.c_str());
			}
		}
	}

	return szProperty;
}

// Caller allocates with new. Clean up with DeleteSPropVal.
_Check_return_ LPSPropValue BinToSPropValue(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, DWORD dwPropCount, _Out_ size_t* lpcbBytesRead, bool bStringPropsExcludeLength)
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

		DWORD dwTemp = 0;
		WORD wTemp = 0;
		switch (PropType)
		{
		case PT_LONG:
			Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.l = dwTemp;
			break;
		case PT_ERROR:
			Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.err = dwTemp;
			break;
		case PT_BOOLEAN:
			Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.b = wTemp;
			break;
		case PT_UNICODE:
			if (bStringPropsExcludeLength)
			{
				Parser.GetStringW(&pspvProperty[i].Value.lpszW);
			}
			else
			{
				// This is apparently a cb...
				Parser.GetWORD(&wTemp);
				Parser.GetStringW(wTemp / sizeof(WCHAR), &pspvProperty[i].Value.lpszW);
			}
			break;
		case PT_STRING8:
			if (bStringPropsExcludeLength)
			{
				Parser.GetStringA(&pspvProperty[i].Value.lpszA);
			}
			else
			{
				// This is apparently a cb...
				Parser.GetWORD(&wTemp);
				Parser.GetStringA(wTemp, &pspvProperty[i].Value.lpszA);
			}
			break;
		case PT_SYSTIME:
			Parser.GetDWORD(&pspvProperty[i].Value.ft.dwHighDateTime);
			Parser.GetDWORD(&pspvProperty[i].Value.ft.dwLowDateTime);
			break;
		case PT_BINARY:
			Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.bin.cb = wTemp;
			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			Parser.GetBYTES(pspvProperty[i].Value.bin.cb, pspvProperty[i].Value.bin.cb, &pspvProperty[i].Value.bin.lpb);
			break;
		case PT_MV_STRING8:
			Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.MVszA.cValues = wTemp;
			pspvProperty[i].Value.MVszA.lppszA = new CHAR*[wTemp];
			if (pspvProperty[i].Value.MVszA.lppszA)
			{
				memset(pspvProperty[i].Value.MVszA.lppszA, 0, sizeof(CHAR*)* wTemp);
				DWORD j = 0;
				for (j = 0; j < pspvProperty[i].Value.MVszA.cValues; j++)
				{
					Parser.GetStringA(&pspvProperty[i].Value.MVszA.lppszA[j]);
				}
			}
			break;
		case PT_MV_UNICODE:
			Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.MVszW.cValues = wTemp;
			pspvProperty[i].Value.MVszW.lppszW = new WCHAR*[wTemp];
			if (pspvProperty[i].Value.MVszW.lppszW)
			{
				memset(pspvProperty[i].Value.MVszW.lppszW, 0, sizeof(WCHAR*)* wTemp);
				DWORD j = 0;
				for (j = 0; j < pspvProperty[i].Value.MVszW.cValues; j++)
				{
					Parser.GetStringW(&pspvProperty[i].Value.MVszW.lppszW[j]);
				}
			}
			break;
		}
	}

	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();
	return pspvProperty;
}

void DeleteSPropVal(ULONG cVal, _In_count_(cVal) LPSPropValue lpsPropVal)
{
	if (!lpsPropVal) return;
	DWORD i = 0;
	DWORD j = 0;
	for (i = 0; i < cVal; i++)
	{
		switch (PROP_TYPE(lpsPropVal[i].ulPropTag))
		{
		case PT_UNICODE:
			delete[] lpsPropVal[i].Value.lpszW;
			break;
		case PT_STRING8:
			delete[] lpsPropVal[i].Value.lpszA;
			break;
		case PT_BINARY:
			delete[] lpsPropVal[i].Value.bin.lpb;
			break;
		case PT_MV_STRING8:
			if (lpsPropVal[i].Value.MVszA.lppszA)
			{
				for (j = 0; j < lpsPropVal[i].Value.MVszA.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVszA.lppszA[j];
				}
			}
			break;
		case PT_MV_UNICODE:
			if (lpsPropVal[i].Value.MVszW.lppszW)
			{
				for (j = 0; j < lpsPropVal[i].Value.MVszW.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVszW.lppszW[j];
				}
			}
			break;
		case PT_MV_BINARY:
			if (lpsPropVal[i].Value.MVbin.lpbin)
			{
				for (j = 0; j < lpsPropVal[i].Value.MVbin.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVbin.lpbin[j].lpb;
				}
			}
			break;
		}
	}

	delete[] lpsPropVal;
}