#include "stdafx.h"
#include "ParseProperty.h"
#include "Property.h"
#include "MAPIFunctions.h"
#include "ExtraPropTags.h"
#include "String.h"

// We avoid bringing InterpretProp.h in with these
void RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj, ULONG ulTabLevel, _In_ CString *PropString);
void ActionsToString(_In_ ACTIONS* lpActions, _In_ CString* PropString);

wstring BuildErrorPropString(_In_ LPSPropValue lpProp)
{
	if (PROP_TYPE(lpProp->ulPropTag) != PT_ERROR) return L"";
	switch (PROP_ID(lpProp->ulPropTag))
	{
	case PROP_ID(PR_BODY):
	case PROP_ID(PR_BODY_HTML):
	case PROP_ID(PR_RTF_COMPRESSED):
		if (MAPI_E_NOT_ENOUGH_MEMORY == lpProp->Value.err ||
			MAPI_E_NOT_FOUND == lpProp->Value.err)
		{
			return loadstring(IDS_OPENBODY);
		}

		break;
	default:
		if (MAPI_E_NOT_ENOUGH_MEMORY == lpProp->Value.err)
		{
			return loadstring(IDS_OPENSTREAM);
		}
	}

	return L"";
}

 wstring BinToTextString(_In_ LPSBinary lpBin, bool bMultiLine)
{
	if (!lpBin || !lpBin->cb || !lpBin->lpb) return L"";

	wstring szBin;

	ULONG i;
	for (i = 0; i < lpBin->cb; i++)
	{
		// Any printable extended ASCII character gets mapped directly
		if (lpBin->lpb[i] >= 0x20 &&
			lpBin->lpb[i] <= 0xFE)
		{
			szBin += lpBin->lpb[i];
		}
		// If we allow multiple lines, we accept tab, LF and CR
		else if (bMultiLine &&
			(lpBin->lpb[i] == 9 || // Tab
			lpBin->lpb[i] == 10 || // Line Feed
			lpBin->lpb[i] == 13))  // Carriage Return
		{
			szBin += lpBin->lpb[i];
		}
		// Everything else is a dot
		else
		{
			szBin += L'.';
		}
	}

	return szBin;
}

 wstring MyHexFromBin(_In_opt_count_(cb) LPBYTE lpb, size_t cb, bool bPrependCB)
{
	wstring lpsz;

	if (bPrependCB)
	{
		lpsz = format(L"cb: %u lpb: ", (UINT)cb); // STRING_OK
	}

	if (!cb || !lpb)
	{
		lpsz += L"NULL";
	}
	else
	{
		ULONG i = 0;
		for (i = 0; i < cb; i++)
		{
			BYTE bLow = (BYTE)((lpb[i]) & 0xf);
			BYTE bHigh = (BYTE)((lpb[i] >> 4) & 0xf);
			wchar_t szLow = (wchar_t)((bLow <= 0x9) ? L'0' + bLow : L'A' + bLow - 0xa);
			wchar_t szHigh = (wchar_t)((bHigh <= 0x9) ? L'0' + bHigh : L'A' + bHigh - 0xa);

			lpsz += szHigh;
			lpsz += szLow;
		}
	}

	return lpsz;
}

 wstring BinToHexString(_In_opt_ LPSBinary lpBin, bool bPrependCB)
{
	if (!lpBin) return L"";

	return MyHexFromBin(
		lpBin->lpb,
		lpBin->cb,
		bPrependCB);
}

 wstring GUIDToWstring(_In_opt_ LPCGUID lpGUID)
{
	GUID nullGUID = { 0 };
	wstring szGUID;

	if (!lpGUID)
	{
		lpGUID = &nullGUID;
	}

	szGUID = format(L"{%.8X-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X}", // STRING_OK
		lpGUID->Data1,
		lpGUID->Data2,
		lpGUID->Data3,
		lpGUID->Data4[0],
		lpGUID->Data4[1],
		lpGUID->Data4[2],
		lpGUID->Data4[3],
		lpGUID->Data4[4],
		lpGUID->Data4[5],
		lpGUID->Data4[6],
		lpGUID->Data4[7]);

	return szGUID;
}

 wstring GUIDToWstringAndName(_In_opt_ LPCGUID lpGUID)
{
	ULONG ulCur = 0;
	wstring szGUIDName;

	if (lpGUID && ulPropGuidArray && PropGuidArray)
	{
		for (ulCur = 0; ulCur < ulPropGuidArray; ulCur++)
		{
			if (IsEqualGUID(*lpGUID, *PropGuidArray[ulCur].lpGuid))
			{
				szGUIDName = PropGuidArray[ulCur].lpszName;
				break;
			}
		}
	}

	if (szGUIDName.empty())
	{
		szGUIDName = loadstring(IDS_UNKNOWNGUID);
	}

	wstring szGUID = GUIDToWstring(lpGUID);
	return szGUIDName + L" = " + szGUID;
}

void FileTimeToString(_In_ FILETIME* lpFileTime, _In_ wstring& PropString, _In_opt_ wstring& AltPropString)
{
	HRESULT hRes = S_OK;
	SYSTEMTIME SysTime = { 0 };

	if (!lpFileTime) return;

	WC_B(FileTimeToSystemTime((FILETIME*)lpFileTime, &SysTime));

	if (S_OK == hRes)
	{
		int iRet = 0;
		wstring szFormatStr;
		wchar_t szTimeStr[MAX_PATH] = { 0 };
		wchar_t szDateStr[MAX_PATH] = { 0 };

		// shove millisecond info into our format string since GetTimeFormat doesn't use it
		szFormatStr = formatmessage(IDS_FILETIMEFORMAT, SysTime.wMilliseconds);

		WC_D(iRet, GetTimeFormatW(
			LOCALE_USER_DEFAULT,
			NULL,
			&SysTime,
			szFormatStr.c_str(),
			szTimeStr,
			MAX_PATH));

		WC_D(iRet, GetDateFormatW(
			LOCALE_USER_DEFAULT,
			NULL,
			&SysTime,
			NULL,
			szDateStr,
			MAX_PATH));

		PropString = format(L"%ws %ws", szTimeStr, szDateStr); // STRING_OK
	}
	else
	{
		PropString = loadstring(IDS_INVALIDSYSTIME);
	}

	AltPropString = formatmessage(IDS_FILETIMEALTFORMAT, lpFileTime->dwLowDateTime, lpFileTime->dwHighDateTime);
}

 wstring RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj)
{
	CString szRes;
	RestrictionToString(lpRes, lpObj, 0, &szRes);

#ifdef UNICODE
	return (LPCTSTR)szRes;
#else
	return format(L"%hs", (LPCTSTR)szRes);
#endif
}

 wstring ActionsToString(_In_ ACTIONS* lpActions)
{
	CString szActions;
	ActionsToString(lpActions, &szActions);

#ifdef UNICODE
	return (LPCTSTR)szActions;
#else
	return format(L"%hs", (LPCTSTR)szActions);
#endif
}

Property ParseMVProperty(_In_ LPSPropValue lpProp, ULONG ulMVRow)
{
	if (!lpProp || ulMVRow > lpProp->Value.MVi.cValues) return Property();

	// We'll let ParseProperty do all the work
	SPropValue sProp = { 0 };
	sProp.ulPropTag = CHANGE_PROP_TYPE(lpProp->ulPropTag, PROP_TYPE(lpProp->ulPropTag) & ~MV_FLAG);

	// Only attempt to dereference our array if it's non-NULL
	if (PROP_TYPE(lpProp->ulPropTag) & MV_FLAG &&
		lpProp->Value.MVi.lpi)
	{
		switch (PROP_TYPE(lpProp->ulPropTag))
		{
		case PT_MV_I2:
			sProp.Value.i = lpProp->Value.MVi.lpi[ulMVRow];
			break;
		case PT_MV_LONG:
			sProp.Value.l = lpProp->Value.MVl.lpl[ulMVRow];
			break;
		case PT_MV_DOUBLE:
			sProp.Value.dbl = lpProp->Value.MVdbl.lpdbl[ulMVRow];
			break;
		case PT_MV_CURRENCY:
			sProp.Value.cur = lpProp->Value.MVcur.lpcur[ulMVRow];
			break;
		case PT_MV_APPTIME:
			sProp.Value.at = lpProp->Value.MVat.lpat[ulMVRow];
			break;
		case PT_MV_SYSTIME:
			sProp.Value.ft = lpProp->Value.MVft.lpft[ulMVRow];
			break;
		case PT_MV_I8:
			sProp.Value.li = lpProp->Value.MVli.lpli[ulMVRow];
			break;
		case PT_MV_R4:
			sProp.Value.flt = lpProp->Value.MVflt.lpflt[ulMVRow];
			break;
		case PT_MV_STRING8:
			sProp.Value.lpszA = lpProp->Value.MVszA.lppszA[ulMVRow];
			break;
		case PT_MV_UNICODE:
			sProp.Value.lpszW = lpProp->Value.MVszW.lppszW[ulMVRow];
			break;
		case PT_MV_BINARY:
			sProp.Value.bin = lpProp->Value.MVbin.lpbin[ulMVRow];
			break;
		case PT_MV_CLSID:
			sProp.Value.lpguid = &lpProp->Value.MVguid.lpguid[ulMVRow];
			break;
		default:
			break;
		}
	}

	return ParseProperty(&sProp);
}

Property ParseProperty(_In_ LPSPropValue lpProp)
{
	Property properties;

	ULONG iMVCount = 0;

	if (!lpProp) return properties;

	if (MV_FLAG & PROP_TYPE(lpProp->ulPropTag))
	{
		// MV property
		properties.AddAttribute(L"mv", L"true"); // STRING_OK
		// All the MV structures are basically the same, so we can cheat when we pull the count
		properties.AddAttribute(L"count", to_wstring(lpProp->Value.MVi.cValues)); // STRING_OK

		// Don't bother with the loop if we don't have data
		if (lpProp->Value.MVi.lpi)
		{
			for (iMVCount = 0; iMVCount < lpProp->Value.MVi.cValues; iMVCount++)
			{
				properties.AddMVParsing(ParseMVProperty(lpProp, iMVCount));
			}
		}
	}
	else
	{
		wstring szTmp;
		bool bPropXMLSafe = true;
		Attributes attributes;

		wstring szAltTmp;
		bool bAltPropXMLSafe = true;
		Attributes altAttributes;

		switch (PROP_TYPE(lpProp->ulPropTag))
		{
		case PT_I2:
			szTmp = format(L"%d", lpProp->Value.i); // STRING_OK
			szAltTmp = format(L"0x%X", lpProp->Value.i); // STRING_OK
			break;
		case PT_LONG:
			szTmp = format(L"%d", lpProp->Value.l); // STRING_OK
			szAltTmp = format(L"0x%X", lpProp->Value.l); // STRING_OK
			break;
		case PT_R4:
			szTmp = format(L"%f", lpProp->Value.flt); // STRING_OK
			break;
		case PT_DOUBLE:
			szTmp = format(L"%f", lpProp->Value.dbl); // STRING_OK
			break;
		case PT_CURRENCY:
			szTmp = format(L"%05I64d", lpProp->Value.cur.int64); // STRING_OK
			if (szTmp.length() > 4)
			{
				szTmp.insert(szTmp.length() - 4, L".");
			}

			szAltTmp = format(L"0x%08X:0x%08X", (int)(lpProp->Value.cur.Hi), (int)lpProp->Value.cur.Lo); // STRING_OK
			break;
		case PT_APPTIME:
			szTmp = format(L"%f", lpProp->Value.at); // STRING_OK
			break;
		case PT_ERROR:
			szTmp = format(L"%ws", ErrorNameFromErrorCode(lpProp->Value.err)); // STRING_OK
			szAltTmp = BuildErrorPropString(lpProp);

			attributes.AddAttribute(L"err", format(L"0x%08X", lpProp->Value.err)); // STRING_OK
			break;
		case PT_BOOLEAN:
			if (lpProp->Value.b)
				szTmp = loadstring(IDS_TRUE);
			else
				szTmp = loadstring(IDS_FALSE);
			break;
		case PT_OBJECT:
			szTmp = loadstring(IDS_OBJECT);
			break;
		case PT_I8: // LARGE_INTEGER
			szTmp = format(L"0x%08X:0x%08X", (int)(lpProp->Value.li.HighPart), (int)lpProp->Value.li.LowPart); // STRING_OK
			szAltTmp = format(L"%I64d", lpProp->Value.li.QuadPart); // STRING_OK
			break;
		case PT_STRING8:
			if (CheckStringProp(lpProp, PT_STRING8))
			{
				szTmp = format(L"%hs", lpProp->Value.lpszA); // STRING_OK
				bPropXMLSafe = false;

				SBinary sBin = { 0 };
				sBin.cb = (ULONG)szTmp.length();
				sBin.lpb = (LPBYTE)lpProp->Value.lpszA;
				szAltTmp = BinToHexString(&sBin, false);

				altAttributes.AddAttribute(L"cb", format(L"%u", sBin.cb)); // STRING_OK
			}
			break;
		case PT_UNICODE:
			if (CheckStringProp(lpProp, PT_UNICODE))
			{
				szTmp = lpProp->Value.lpszW;
				bPropXMLSafe = false;

				SBinary sBin = { 0 };
				sBin.cb = ((ULONG)szTmp.length()) * sizeof(WCHAR);
				sBin.lpb = (LPBYTE)lpProp->Value.lpszW;
				szAltTmp = BinToHexString(&sBin, false);

				altAttributes.AddAttribute(L"cb", format(L"%u", sBin.cb)); // STRING_OK
			}
			break;
		case PT_SYSTIME:
			FileTimeToString(&lpProp->Value.ft, szTmp, szAltTmp);
			break;
		case PT_CLSID:
			// TODO: One string matches current behavior - look at splitting to two strings in future change
			szTmp = GUIDToWstringAndName(lpProp->Value.lpguid);
			break;
		case PT_BINARY:
			szTmp = BinToHexString(&lpProp->Value.bin, false);
			szAltTmp = BinToTextString(&lpProp->Value.bin, true);
			bAltPropXMLSafe = false;

			attributes.AddAttribute(L"cb", format(L"%u", lpProp->Value.bin.cb)); // STRING_OK
			break;
		case PT_SRESTRICTION:
			szTmp = RestrictionToString((LPSRestriction)lpProp->Value.lpszA, NULL);
			bPropXMLSafe = false;
			break;
		case PT_ACTIONS:
			szTmp = ActionsToString((ACTIONS*)lpProp->Value.lpszA);
			bPropXMLSafe = false;
			break;
		default:
			break;
		}

		Parsing mainParsing(szTmp, bPropXMLSafe, attributes);
		Parsing altParsing(szAltTmp, bAltPropXMLSafe, altAttributes);
		properties.AddParsing(mainParsing, altParsing);
	}

	return properties;
}