#include "stdafx.h"
#include "ParseProperty.h"
#include "Property.h"
#include "MAPIFunctions.h"
#include "ExtraPropTags.h"
#include "String.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include <cctype>

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
			for (ULONG iMVCount = 0; iMVCount < lpProp->Value.MVi.cValues; iMVCount++)
			{
				properties.AddMVParsing(ParseMVProperty(lpProp, iMVCount));
			}
		}
	}
	else
	{
		wstring szTmp;
		auto bPropXMLSafe = true;
		Attributes attributes;

		wstring szAltTmp;
		auto bAltPropXMLSafe = true;
		Attributes altAttributes;

		switch (PROP_TYPE(lpProp->ulPropTag))
		{
		case PT_I2:
			szTmp = to_wstring(lpProp->Value.i);
			szAltTmp = format(L"0x%X", lpProp->Value.i); // STRING_OK
			break;
		case PT_LONG:
			szTmp = to_wstring(lpProp->Value.l);
			szAltTmp = format(L"0x%X", lpProp->Value.l); // STRING_OK
			break;
		case PT_R4:
			szTmp = to_wstring(lpProp->Value.flt); // STRING_OK
			break;
		case PT_DOUBLE:
			szTmp = to_wstring(lpProp->Value.dbl); // STRING_OK
			break;
		case PT_CURRENCY:
			szTmp = format(L"%05I64d", lpProp->Value.cur.int64); // STRING_OK
			if (szTmp.length() > 4)
			{
				szTmp.insert(szTmp.length() - 4, L".");
			}

			szAltTmp = format(L"0x%08X:0x%08X", static_cast<int>(lpProp->Value.cur.Hi), static_cast<int>(lpProp->Value.cur.Lo)); // STRING_OK
			break;
		case PT_APPTIME:
			szTmp = to_wstring(lpProp->Value.at); // STRING_OK
			break;
		case PT_ERROR:
			szTmp = ErrorNameFromErrorCode(lpProp->Value.err); // STRING_OK
			szAltTmp = BuildErrorPropString(lpProp);

			attributes.AddAttribute(L"err", format(L"0x%08X", lpProp->Value.err)); // STRING_OK
			break;
		case PT_BOOLEAN:
			szTmp = loadstring(lpProp->Value.b ? IDS_TRUE : IDS_FALSE);
			break;
		case PT_OBJECT:
			szTmp = loadstring(IDS_OBJECT);
			break;
		case PT_I8: // LARGE_INTEGER
			szTmp = format(L"0x%08X:0x%08X", static_cast<int>(lpProp->Value.li.HighPart), static_cast<int>(lpProp->Value.li.LowPart)); // STRING_OK
			szAltTmp = format(L"%I64d", lpProp->Value.li.QuadPart); // STRING_OK
			break;
		case PT_STRING8:
			if (CheckStringProp(lpProp, PT_STRING8))
			{
				szTmp = LPCSTRToWstring(lpProp->Value.lpszA);
				bPropXMLSafe = false;

				SBinary sBin = { 0 };
				sBin.cb = static_cast<ULONG>(szTmp.length());
				sBin.lpb = reinterpret_cast<LPBYTE>(lpProp->Value.lpszA);
				szAltTmp = BinToHexString(&sBin, false);

				altAttributes.AddAttribute(L"cb", to_wstring(sBin.cb)); // STRING_OK
			}
			break;
		case PT_UNICODE:
			if (CheckStringProp(lpProp, PT_UNICODE))
			{
				szTmp = lpProp->Value.lpszW;
				bPropXMLSafe = false;

				SBinary sBin = { 0 };
				sBin.cb = static_cast<ULONG>(szTmp.length()) * sizeof(WCHAR);
				sBin.lpb = reinterpret_cast<LPBYTE>(lpProp->Value.lpszW);
				szAltTmp = BinToHexString(&sBin, false);

				altAttributes.AddAttribute(L"cb", to_wstring(sBin.cb)); // STRING_OK
			}
			break;
		case PT_SYSTIME:
			FileTimeToString(&lpProp->Value.ft, szTmp, szAltTmp);
			break;
		case PT_CLSID:
			// TODO: One string matches current behavior - look at splitting to two strings in future change
			szTmp = GUIDToStringAndName(lpProp->Value.lpguid);
			break;
		case PT_BINARY:
			szTmp = BinToHexString(&lpProp->Value.bin, false);
			szAltTmp = BinToTextString(&lpProp->Value.bin, true);
			bAltPropXMLSafe = false;

			attributes.AddAttribute(L"cb", to_wstring(lpProp->Value.bin.cb)); // STRING_OK
			break;
		case PT_SRESTRICTION:
			szTmp = RestrictionToString(reinterpret_cast<LPSRestriction>(lpProp->Value.lpszA), nullptr);
			bPropXMLSafe = false;
			break;
		case PT_ACTIONS:
			szTmp = ActionsToString(reinterpret_cast<ACTIONS*>(lpProp->Value.lpszA));
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