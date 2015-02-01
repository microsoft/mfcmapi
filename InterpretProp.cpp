#include "stdafx.h"
#include "InterpretProp.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"
#include "NamedPropCache.h"
#include "SmartView\SmartView.h"
#include "ParseProperty.h"
#include "String.h"

static const char pBase64[] = {
	0x3e, 0x7f, 0x7f, 0x7f, 0x3f, 0x34, 0x35, 0x36,
	0x37, 0x38, 0x39, 0x3a, 0x3b, 0x3c, 0x3d, 0x7f,
	0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x00, 0x01,
	0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09,
	0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10, 0x11,
	0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19,
	0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x1a, 0x1b,
	0x1c, 0x1d, 0x1e, 0x1f, 0x20, 0x21, 0x22, 0x23,
	0x24, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2a, 0x2b,
	0x2c, 0x2d, 0x2e, 0x2f, 0x30, 0x31, 0x32, 0x33
};

// allocates output buffer with new
// delete with delete[]
// suprisingly, this algorithm works in a unicode build as well
_Check_return_ HRESULT Base64Decode(_In_z_ LPCTSTR szEncodedStr, _Inout_ size_t* cbBuf, _Out_ _Deref_post_cap_(*cbBuf) LPBYTE* lpDecodedBuffer)
{
	HRESULT hRes = S_OK;
	size_t	cchLen = 0;

	EC_H(StringCchLength(szEncodedStr, STRSAFE_MAX_CCH, &cchLen));

	if (cchLen % 4) return MAPI_E_INVALID_PARAMETER;

	// look for padding at the end
	static const TCHAR szPadding[] = _T("=="); // STRING_OK
	const TCHAR* szPaddingLoc = NULL;
	szPaddingLoc = _tcschr(szEncodedStr, szPadding[0]);
	size_t cchPaddingLen = 0;
	if (NULL != szPaddingLoc)
	{
		// check padding length
		EC_H(StringCchLength(szPaddingLoc, STRSAFE_MAX_CCH, &cchPaddingLen));
		if (cchPaddingLen >= 3) return MAPI_E_INVALID_PARAMETER;

		// check for bad characters after the first '='
		if (_tcsncmp(szPaddingLoc, (TCHAR *)szPadding, cchPaddingLen)) return MAPI_E_INVALID_PARAMETER;
	}
	// cchPaddingLen == 0,1,2 now

	size_t	cchDecodedLen = ((cchLen + 3) / 4) * 3; // 3 times number of 4 tuplets, rounded up

	// back off the decoded length to the correct length
	// xx== ->y
	// xxx= ->yY
	// x=== ->this is a bad case which should never happen
	cchDecodedLen -= cchPaddingLen;
	// we have no room for error now!
	*lpDecodedBuffer = new BYTE[cchDecodedLen];
	if (!*lpDecodedBuffer) return MAPI_E_CALL_FAILED;

	*cbBuf = cchDecodedLen;

	LPBYTE	lpOutByte = *lpDecodedBuffer;

	TCHAR c[4] = { 0 };
	BYTE bTmp[3] = { 0 }; // output

	while (*szEncodedStr)
	{
		int i = 0;
		int iOutlen = 3;
		for (i = 0; i < 4; i++)
		{
			c[i] = *(szEncodedStr + i);
			if (c[i] == _T('='))
			{
				iOutlen = i - 1;
				break;
			}
			if ((c[i] < 0x2b) || (c[i] > 0x7a)) return MAPI_E_INVALID_PARAMETER;

			c[i] = pBase64[c[i] - 0x2b];
		}
		bTmp[0] = (BYTE)((c[0] << 2) | (c[1] >> 4));
		bTmp[1] = (BYTE)((c[1] & 0x0f) << 4 | (c[2] >> 2));
		bTmp[2] = (BYTE)((c[2] & 0x03) << 6 | c[3]);

		for (i = 0; i < iOutlen; i++)
		{
			lpOutByte[i] = bTmp[i];
		}
		lpOutByte += 3;
		szEncodedStr += 4;
	}

	return hRes;
} // Base64Decode

static const		// Base64 Index into encoding
char pIndex[] = {	// and decoding table.
	0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48,
	0x49, 0x4a, 0x4b, 0x4c, 0x4d, 0x4e, 0x4f, 0x50,
	0x51, 0x52, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58,
	0x59, 0x5a, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66,
	0x67, 0x68, 0x69, 0x6a, 0x6b, 0x6c, 0x6d, 0x6e,
	0x6f, 0x70, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76,
	0x77, 0x78, 0x79, 0x7a, 0x30, 0x31, 0x32, 0x33,
	0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x2b, 0x2f
};

// allocates output string with new
// delete with delete[]
_Check_return_ HRESULT Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) LPBYTE lpSourceBuffer, _Inout_ size_t* cchEncodedStr, _Out_ _Deref_post_cap_(*cchEncodedStr) LPTSTR* szEncodedStr)
{
	HRESULT hRes = S_OK;

	size_t cchEncodeLen = ((cbSourceBuf + 2) / 3) * 4; // 4 * number of size three blocks, round up, plus null terminator
	*szEncodedStr = new TCHAR[cchEncodeLen + 1]; // allocate a touch extra for some NULL terminators
	if (cchEncodedStr) *cchEncodedStr = cchEncodeLen;
	if (!*szEncodedStr) return MAPI_E_CALL_FAILED;

	size_t cbBuf = 0; // General purpose integers.
	TCHAR* szOutChar = NULL;
	szOutChar = *szEncodedStr;

	// Using integer division to round down here
	while (cbBuf < (cbSourceBuf / 3) * 3) // encode each 3 byte octet.
	{
		*szOutChar++ = pIndex[lpSourceBuffer[cbBuf] >> 2];
		*szOutChar++ = pIndex[((lpSourceBuffer[cbBuf] & 0x03) << 4) + (lpSourceBuffer[cbBuf + 1] >> 4)];
		*szOutChar++ = pIndex[((lpSourceBuffer[cbBuf + 1] & 0x0f) << 2) + (lpSourceBuffer[cbBuf + 2] >> 6)];
		*szOutChar++ = pIndex[lpSourceBuffer[cbBuf + 2] & 0x3f];
		cbBuf += 3; // Next octet.
	}

	if (cbSourceBuf - cbBuf) // Partial octet remaining?
	{
		*szOutChar++ = pIndex[lpSourceBuffer[cbBuf] >> 2]; // Yes, encode it.

		if (cbSourceBuf - cbBuf == 1) // End of octet?
		{
			*szOutChar++ = pIndex[(lpSourceBuffer[cbBuf] & 0x03) << 4];
			*szOutChar++ = _T('=');
			*szOutChar++ = _T('=');
		}
		else
		{ // No, one more part.
			*szOutChar++ = pIndex[((lpSourceBuffer[cbBuf] & 0x03) << 4) + (lpSourceBuffer[cbBuf + 1] >> 4)];
			*szOutChar++ = pIndex[(lpSourceBuffer[cbBuf + 1] & 0x0f) << 2];
			*szOutChar++ = _T('=');
		}
	}
	*szOutChar = _T('\0');

	return hRes;
} // Base64Encode

// Allocates string for GUID with new
// free with delete[]
#define GUID_STRING_SIZE 39
_Check_return_ LPTSTR GUIDToString(_In_opt_ LPCGUID lpGUID)
{
	HRESULT	hRes = S_OK;
	GUID	nullGUID = { 0 };
	LPTSTR	szGUID = NULL;

	if (!lpGUID)
	{
		lpGUID = &nullGUID;
	}

	szGUID = new TCHAR[GUID_STRING_SIZE];

	EC_H(StringCchPrintf(szGUID, GUID_STRING_SIZE, _T("{%.8X-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X}"), // STRING_OK
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
		lpGUID->Data4[7]));

	return szGUID;
} // GUIDToString

_Check_return_ HRESULT StringToGUID(_In_z_ LPCTSTR szGUID, _Inout_ LPGUID lpGUID)
{
	return StringToGUID(szGUID, false, lpGUID);
} // StringToGUID

_Check_return_ HRESULT StringToGUID(_In_z_ LPCTSTR szGUID, bool bByteSwapped, _Inout_ LPGUID lpGUID)
{
	HRESULT hRes = S_OK;
	if (!szGUID || !lpGUID) return MAPI_E_INVALID_PARAMETER;

	ULONG cbGUID = sizeof(GUID);

	// Now we use MyBinFromHex to do the work.
	(void)MyBinFromHex(szGUID, (LPBYTE)lpGUID, &cbGUID);

	// Note that we get the bByteSwapped behavior by default. We have to work to get the 'normal' behavior
	if (!bByteSwapped)
	{
		LPBYTE lpByte = (LPBYTE)lpGUID;
		BYTE bByte = 0;
		bByte = lpByte[0];
		lpByte[0] = lpByte[3];
		lpByte[3] = bByte;
		bByte = lpByte[1];
		lpByte[1] = lpByte[2];
		lpByte[2] = bByte;
	}
	return hRes;
} // StringToGUID

_Check_return_ CString CurrencyToString(CURRENCY curVal)
{
	CString szCur;

	szCur.Format(_T("%05I64d"), curVal.int64); // STRING_OK
	if (szCur.GetLength() > 4)
	{
		szCur.Insert(szCur.GetLength() - 4, _T(".")); // STRING_OK
	}
	return szCur;
} // CurrencyToString

_Check_return_ CString TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine)
{
	CString szRet;
	CString szTemp;
	HRESULT hRes = S_OK;

	LPTSTR szExactMatches = NULL;
	LPTSTR szPartialMatches = NULL;
	LPTSTR szNamedPropName = NULL;
	LPTSTR szNamedPropGUID = NULL;
	LPTSTR szNamedPropDASL = NULL;

	NameIDToStrings(
		ulPropTag,
		lpObj,
		NULL,
		NULL,
		bIsAB,
		&szNamedPropName, // Built from lpProp & lpMAPIProp
		&szNamedPropGUID, // Built from lpProp & lpMAPIProp
		&szNamedPropDASL); // Built from ulPropTag & lpMAPIProp

	PropTagToPropName(ulPropTag, bIsAB, &szExactMatches, &szPartialMatches);

	CString szFormatString;
	if (bSingleLine)
	{
		szFormatString = _T("0x%1!08X! (%2)"); // STRING_OK
		if (!IsNullOrEmpty(szExactMatches)) szFormatString += _T(": %3"); // STRING_OK
		if (!IsNullOrEmpty(szPartialMatches)) szFormatString += _T(": (%4)"); // STRING_OK
		if (szNamedPropName)
		{
			EC_B(szTemp.LoadString(IDS_NAMEDPROPSINGLELINE));
			szFormatString += szTemp;
		}
		if (szNamedPropGUID)
		{
			EC_B(szTemp.LoadString(IDS_GUIDSINGLELINE));
			szFormatString += szTemp;
		}
	}
	else
	{
		EC_B(szFormatString.LoadString(IDS_TAGMULTILINE));
		if (!IsNullOrEmpty(szExactMatches))
		{
			EC_B(szTemp.LoadString(IDS_PROPNAMEMULTILINE));
			szFormatString += szTemp;
		}
		if (!IsNullOrEmpty(szPartialMatches))
		{
			EC_B(szTemp.LoadString(IDS_OTHERNAMESMULTILINE));
			szFormatString += szTemp;
		}
		if (PROP_ID(ulPropTag) < 0x8000)
		{
			EC_B(szTemp.LoadString(IDS_DASLPROPTAG));
			szFormatString += szTemp;
		}
		else if (szNamedPropDASL)
		{
			EC_B(szTemp.LoadString(IDS_DASLNAMED));
			szFormatString += szTemp;
		}
		if (szNamedPropName)
		{
			EC_B(szTemp.LoadString(IDS_NAMEPROPNAMEMULTILINE));
			szFormatString += szTemp;
		}
		if (szNamedPropGUID)
		{
			EC_B(szTemp.LoadString(IDS_NAMEPROPGUIDMULTILINE));
			szFormatString += szTemp;
		}
	}
	szRet.FormatMessage(szFormatString,
		ulPropTag,
		(LPCTSTR)TypeToString(ulPropTag),
		szExactMatches,
		szPartialMatches,
		szNamedPropName,
		szNamedPropGUID,
		szNamedPropDASL);

	delete[] szPartialMatches;
	delete[] szExactMatches;
	FreeNameIDStrings(szNamedPropName, szNamedPropGUID, szNamedPropDASL);

	if (fIsSet(DBGTest))
	{
		static size_t cchMaxBuff = 0;
		size_t cchBuff = szRet.GetLength();
		cchMaxBuff = max(cchBuff, cchMaxBuff);
		DebugPrint(DBGTest, _T("TagToString parsing 0x%08X returned %u chars - max %u\n"), ulPropTag, (UINT)cchBuff, (UINT)cchMaxBuff);
	}
	return szRet;
} // TagToString

_Check_return_ CString ProblemArrayToString(_In_ LPSPropProblemArray lpProblems)
{
	CString szOut;
	if (lpProblems)
	{
		ULONG i = 0;
		for (i = 0; i < lpProblems->cProblem; i++)
		{
			CString szTemp;
			szTemp.FormatMessage(IDS_PROBLEMARRAY,
				lpProblems->aProblem[i].ulIndex,
				TagToString(lpProblems->aProblem[i].ulPropTag, NULL, false, false),
				lpProblems->aProblem[i].scode,
				ErrorNameFromErrorCode(lpProblems->aProblem[i].scode));
			szOut += szTemp;
		}
	}
	return szOut;
} // ProblemArrayToString

_Check_return_ CString MAPIErrToString(ULONG ulFlags, _In_ LPMAPIERROR lpErr)
{
	CString szOut;
	if (lpErr)
	{
		szOut.FormatMessage(
			ulFlags & MAPI_UNICODE ? IDS_MAPIERRUNICODE : IDS_MAPIERRANSI,
			lpErr->ulVersion,
			lpErr->lpszError,
			lpErr->lpszComponent,
			lpErr->ulLowLevelError,
			ErrorNameFromErrorCode(lpErr->ulLowLevelError),
			lpErr->ulContext);
	}
	return szOut;
} // MAPIErrToString

_Check_return_ CString TnefProblemArrayToString(_In_ LPSTnefProblemArray lpError)
{
	CString szOut;
	if (lpError)
	{
		for (ULONG iError = 0; iError < lpError->cProblem; iError++)
		{
			CString szTemp;
			szTemp.FormatMessage(
				IDS_TNEFPROBARRAY,
				lpError->aProblem[iError].ulComponent,
				lpError->aProblem[iError].ulAttribute,
				TagToString(lpError->aProblem[iError].ulPropTag, NULL, false, false),
				lpError->aProblem[iError].scode,
				ErrorNameFromErrorCode(lpError->aProblem[iError].scode));
			szOut += szTemp;
		}
	}
	return szOut;
} // TnefProblemArrayToString

// There may be restrictions with over 100 nested levels, but we're not going to try to parse them
#define _MaxRestrictionNesting 100

void RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj, ULONG ulTabLevel, _In_ CString *PropString)
{
	if (!PropString) return;

	*PropString = _T(""); // STRING_OK

	ULONG i = 0;
	if (!lpRes)
	{
		PropString->FormatMessage(IDS_NULLRES);
		return;
	}
	if (ulTabLevel > _MaxRestrictionNesting)
	{
		PropString->FormatMessage(IDS_RESDEPTHEXCEEDED);
		return;
	}
	CString szTmp;
	wstring szProp;
	wstring szAltProp;

	CString szTabs;
	for (i = 0; i < ulTabLevel; i++)
	{
		szTabs += _T("\t"); // STRING_OK
	}

	LPTSTR szFlags = NULL;
	wstring szPropNum;
	InterpretFlags(flagRestrictionType, lpRes->rt, &szFlags);
	szTmp.FormatMessage(IDS_RESTYPE, szTabs, lpRes->rt, szFlags);
	*PropString += szTmp;
	delete[] szFlags;
	szFlags = NULL;
	switch (lpRes->rt)
	{
	case RES_COMPAREPROPS:
		InterpretFlags(flagRelop, lpRes->res.resCompareProps.relop, &szFlags);
		szTmp.FormatMessage(
			IDS_RESCOMPARE,
			szTabs,
			szFlags,
			lpRes->res.resCompareProps.relop,
			TagToString(lpRes->res.resCompareProps.ulPropTag1, lpObj, false, true),
			TagToString(lpRes->res.resCompareProps.ulPropTag2, lpObj, false, true));
		*PropString += szTmp;
		delete[] szFlags;
		break;
	case RES_AND:
		szTmp.FormatMessage(IDS_RESANDCOUNT, szTabs, lpRes->res.resAnd.cRes);
		*PropString += szTmp;
		if (lpRes->res.resAnd.lpRes)
		{
			for (i = 0; i < lpRes->res.resAnd.cRes; i++)
			{
				szTmp.FormatMessage(IDS_RESANDPOINTER, szTabs, i);
				*PropString += szTmp;
				RestrictionToString(&lpRes->res.resAnd.lpRes[i], lpObj, ulTabLevel + 1, &szTmp);
				*PropString += szTmp;
			}
		}
		break;
	case RES_OR:
		szTmp.FormatMessage(IDS_RESORCOUNT, szTabs, lpRes->res.resOr.cRes);
		*PropString += szTmp;
		if (lpRes->res.resOr.lpRes)
		{
			for (i = 0; i < lpRes->res.resOr.cRes; i++)
			{
				szTmp.FormatMessage(IDS_RESORPOINTER, szTabs, i);
				*PropString += szTmp;
				RestrictionToString(&lpRes->res.resOr.lpRes[i], lpObj, ulTabLevel + 1, &szTmp);
				*PropString += szTmp;
			}
		}
		break;
	case RES_NOT:
		szTmp.FormatMessage(
			IDS_RESNOT,
			szTabs,
			lpRes->res.resNot.ulReserved);
		*PropString += szTmp;
		RestrictionToString(lpRes->res.resNot.lpRes, lpObj, ulTabLevel + 1, &szTmp);
		*PropString += szTmp;
		break;
	case RES_COUNT:
		// RES_COUNT and RES_NOT look the same, so we use the resNot member here
		szTmp.FormatMessage(
			IDS_RESCOUNT,
			szTabs,
			lpRes->res.resNot.ulReserved);
		*PropString += szTmp;
		RestrictionToString(lpRes->res.resNot.lpRes, lpObj, ulTabLevel + 1, &szTmp);
		*PropString += szTmp;
		break;
	case RES_CONTENT:
		InterpretFlags(flagFuzzyLevel, lpRes->res.resContent.ulFuzzyLevel, &szFlags);
		szTmp.FormatMessage(
			IDS_RESCONTENT,
			szTabs,
			szFlags,
			lpRes->res.resContent.ulFuzzyLevel,
			TagToString(lpRes->res.resContent.ulPropTag, lpObj, false, true));
		delete[] szFlags;
		szFlags = NULL;
		*PropString += szTmp;
		if (lpRes->res.resContent.lpProp)
		{
			InterpretProp(lpRes->res.resContent.lpProp, &szProp, &szAltProp);
			szTmp.FormatMessage(
				IDS_RESCONTENTPROP,
				szTabs,
				TagToString(lpRes->res.resContent.lpProp->ulPropTag, lpObj, false, true),
				szProp.c_str(),
				szAltProp.c_str());
			*PropString += szTmp;
		}
		break;
	case RES_PROPERTY:
		InterpretFlags(flagRelop, lpRes->res.resProperty.relop, &szFlags);
		szTmp.FormatMessage(
			IDS_RESPROP,
			szTabs,
			szFlags,
			lpRes->res.resProperty.relop,
			TagToString(lpRes->res.resProperty.ulPropTag, lpObj, false, true));
		delete[] szFlags;
		szFlags = NULL;
		*PropString += szTmp;
		if (lpRes->res.resProperty.lpProp)
		{
			InterpretProp(lpRes->res.resProperty.lpProp, &szProp, &szAltProp);
			szTmp.FormatMessage(
				IDS_RESPROPPROP,
				szTabs,
				TagToString(lpRes->res.resProperty.lpProp->ulPropTag, lpObj, false, true),
				szProp.c_str(),
				szAltProp.c_str());
			*PropString += szTmp;
			szPropNum = InterpretNumberAsString(lpRes->res.resProperty.lpProp->Value, lpRes->res.resProperty.lpProp->ulPropTag, NULL, NULL, NULL, false);
			if (!szPropNum.empty())
			{
				szTmp.FormatMessage(IDS_RESPROPPROPFLAGS, szTabs, szPropNum.c_str());
				*PropString += szTmp;
			}
		}
		break;
	case RES_BITMASK:
		InterpretFlags(flagBitmask, lpRes->res.resBitMask.relBMR, &szFlags);
		szTmp.FormatMessage(
			IDS_RESBITMASK,
			szTabs,
			szFlags,
			lpRes->res.resBitMask.relBMR,
			lpRes->res.resBitMask.ulMask);
		delete[] szFlags;
		szFlags = NULL;
		*PropString += szTmp;
		szPropNum = InterpretNumberAsStringProp(lpRes->res.resBitMask.ulMask, lpRes->res.resBitMask.ulPropTag);
		if (!szPropNum.empty())
		{
			szTmp.FormatMessage(IDS_RESBITMASKFLAGS, szPropNum.c_str());
			*PropString += szTmp;
		}
		szTmp.FormatMessage(
			IDS_RESBITMASKTAG,
			szTabs,
			TagToString(lpRes->res.resBitMask.ulPropTag, lpObj, false, true));
		*PropString += szTmp;
		break;
	case RES_SIZE:
		InterpretFlags(flagRelop, lpRes->res.resSize.relop, &szFlags);
		szTmp.FormatMessage(
			IDS_RESSIZE,
			szTabs,
			szFlags,
			lpRes->res.resSize.relop,
			lpRes->res.resSize.cb,
			TagToString(lpRes->res.resSize.ulPropTag, lpObj, false, true));
		delete[] szFlags;
		szFlags = NULL;
		*PropString += szTmp;
		break;
	case RES_EXIST:
		szTmp.FormatMessage(
			IDS_RESEXIST,
			szTabs,
			TagToString(lpRes->res.resExist.ulPropTag, lpObj, false, true),
			lpRes->res.resExist.ulReserved1,
			lpRes->res.resExist.ulReserved2);
		*PropString += szTmp;
		break;
	case RES_SUBRESTRICTION:
		szTmp.FormatMessage(
			IDS_RESSUBRES,
			szTabs,
			TagToString(lpRes->res.resSub.ulSubObject, lpObj, false, true));
		*PropString += szTmp;
		RestrictionToString(lpRes->res.resSub.lpRes, lpObj, ulTabLevel + 1, &szTmp);
		*PropString += szTmp;
		break;
	case RES_COMMENT:
		szTmp.FormatMessage(IDS_RESCOMMENT, szTabs, lpRes->res.resComment.cValues);
		*PropString += szTmp;
		if (lpRes->res.resComment.lpProp)
		{
			for (i = 0; i < lpRes->res.resComment.cValues; i++)
			{
				InterpretProp(&lpRes->res.resComment.lpProp[i], &szProp, &szAltProp);
				szTmp.FormatMessage(
					IDS_RESCOMMENTPROPS,
					szTabs,
					i,
					TagToString(lpRes->res.resComment.lpProp[i].ulPropTag, lpObj, false, true),
					szProp.c_str(),
					szAltProp.c_str());
				*PropString += szTmp;
			}
		}
		szTmp.FormatMessage(
			IDS_RESCOMMENTRES,
			szTabs);
		*PropString += szTmp;
		RestrictionToString(lpRes->res.resComment.lpRes, lpObj, ulTabLevel + 1, &szTmp);
		*PropString += szTmp;
		break;
	case RES_ANNOTATION:
		szTmp.FormatMessage(IDS_RESANNOTATION, szTabs, lpRes->res.resComment.cValues);
		*PropString += szTmp;
		if (lpRes->res.resComment.lpProp)
		{
			for (i = 0; i < lpRes->res.resComment.cValues; i++)
			{
				InterpretProp(&lpRes->res.resComment.lpProp[i], &szProp, &szAltProp);
				szTmp.FormatMessage(
					IDS_RESANNOTATIONPROPS,
					szTabs,
					i,
					TagToString(lpRes->res.resComment.lpProp[i].ulPropTag, lpObj, false, true),
					szProp.c_str(),
					szAltProp.c_str());
				*PropString += szTmp;
			}
		}
		szTmp.FormatMessage(
			IDS_RESANNOTATIONRES,
			szTabs);
		*PropString += szTmp;
		RestrictionToString(lpRes->res.resComment.lpRes, lpObj, ulTabLevel + 1, &szTmp);
		*PropString += szTmp;
		break;
	}
} // RestrictionToString

_Check_return_ CString RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj)
{
	CString szRes;
	RestrictionToString(lpRes, lpObj, 0, &szRes);
	return szRes;
} // RestrictionToString

void AdrListToString(_In_ LPADRLIST lpAdrList, _In_ wstring *PropString)
{
	if (!PropString) return;

	*PropString = L""; // STRING_OK
	if (!lpAdrList)
	{
		*PropString = formatmessage(IDS_ADRLISTNULL);
		return;
	}

	wstring szTmp;
	wstring szProp;
	wstring szAltProp;
	*PropString = formatmessage(IDS_ADRLISTCOUNT, lpAdrList->cEntries);

	ULONG i = 0;
	for (i = 0; i < lpAdrList->cEntries; i++)
	{
		szTmp = formatmessage(IDS_ADRLISTENTRIESCOUNT, i, lpAdrList->aEntries[i].cValues);
		*PropString += szTmp;

		ULONG j = 0;
		for (j = 0; j < lpAdrList->aEntries[i].cValues; j++)
		{
			InterpretProp(&lpAdrList->aEntries[i].rgPropVals[j], &szProp, &szAltProp);
			szTmp = formatmessage(
				IDS_ADRLISTENTRY,
				i,
				j,
				TagToString(lpAdrList->aEntries[i].rgPropVals[j].ulPropTag, NULL, false, false),
				szProp.c_str(),
				szAltProp.c_str());
			*PropString += szTmp;
		}
	}
}

void ActionToString(_In_ ACTION* lpAction, _In_ CString* PropString)
{
	if (!PropString) return;

	*PropString = _T(""); // STRING_OK
	if (!lpAction)
	{
		PropString->FormatMessage(IDS_ACTIONNULL);
		return;
	}
	CString szTmp;
	wstring szProp;
	wstring szAltProp;
	LPTSTR szFlags = NULL;
	LPTSTR szFlags2 = NULL;
	InterpretFlags(flagAccountType, lpAction->acttype, &szFlags);
	InterpretFlags(flagRuleFlag, lpAction->ulFlags, &szFlags2);
	PropString->FormatMessage(
		IDS_ACTION,
		lpAction->acttype,
		szFlags,
		RestrictionToString(lpAction->lpRes, NULL),
		lpAction->ulFlags,
		szFlags2);
	delete[] szFlags2;
	delete[] szFlags;
	szFlags2 = NULL;
	szFlags = NULL;

	switch (lpAction->acttype)
	{
	case OP_MOVE:
	case OP_COPY:
	{
					SBinary sBinStore = { 0 };
					SBinary sBinFld = { 0 };
					sBinStore.cb = lpAction->actMoveCopy.cbStoreEntryId;
					sBinStore.lpb = (LPBYTE)lpAction->actMoveCopy.lpStoreEntryId;
					sBinFld.cb = lpAction->actMoveCopy.cbFldEntryId;
					sBinFld.lpb = (LPBYTE)lpAction->actMoveCopy.lpFldEntryId;

					szTmp.FormatMessage(IDS_ACTIONOPMOVECOPY,
						BinToHexString(&sBinStore, true).c_str(),
						BinToTextString(&sBinStore, false).c_str(),
						BinToHexString(&sBinFld, true).c_str(),
						BinToTextString(&sBinFld, false).c_str());
					*PropString += szTmp;
					break;
	}
	case OP_REPLY:
	case OP_OOF_REPLY:
	{

						 SBinary sBin = { 0 };
						 sBin.cb = lpAction->actReply.cbEntryId;
						 sBin.lpb = (LPBYTE)lpAction->actReply.lpEntryId;
						 wstring szGUID = GUIDToStringAndName(&lpAction->actReply.guidReplyTemplate);

						 szTmp.FormatMessage(IDS_ACTIONOPREPLY,
							 BinToHexString(&sBin, true).c_str(),
							 BinToTextString(&sBin, false).c_str(),
							 szGUID.c_str());
						 *PropString += szTmp;
						 break;
	}
	case OP_DEFER_ACTION:
	{
							SBinary sBin = { 0 };
							sBin.cb = lpAction->actDeferAction.cbData;
							sBin.lpb = (LPBYTE)lpAction->actDeferAction.pbData;

							szTmp.FormatMessage(IDS_ACTIONOPDEFER,
								BinToHexString(&sBin, true).c_str(),
								BinToTextString(&sBin, false).c_str());
							*PropString += szTmp;
							break;
	}
	case OP_BOUNCE:
	{
					  InterpretFlags(flagBounceCode, lpAction->scBounceCode, &szFlags);
					  szTmp.FormatMessage(IDS_ACTIONOPBOUNCE, lpAction->scBounceCode, szFlags);
					  delete[] szFlags;
					  szFlags = NULL;
					  *PropString += szTmp;
					  break;
	}
	case OP_FORWARD:
	case OP_DELEGATE:
	{
						szTmp.FormatMessage(IDS_ACTIONOPFORWARDDEL);
						*PropString += szTmp;
						AdrListToString(lpAction->lpadrlist, &szProp);
						*PropString += wstringToCString(szProp);
						break;
	}

	case OP_TAG:
	{
				   InterpretProp(&lpAction->propTag, &szProp, &szAltProp);
				   szTmp.FormatMessage(IDS_ACTIONOPTAG,
					   TagToString(lpAction->propTag.ulPropTag, NULL, false, true),
					   szProp.c_str(),
					   szAltProp.c_str());
				   *PropString += szTmp;
				   break;
	}
	}

	switch (lpAction->acttype)
	{
	case OP_REPLY:
	{
					 InterpretFlags(flagOPReply, lpAction->ulActionFlavor, &szFlags);
					 break;
	}
	case OP_FORWARD:
	{
					   InterpretFlags(flagOpForward, lpAction->ulActionFlavor, &szFlags);
					   break;
	}
	}
	szTmp.FormatMessage(IDS_ACTIONFLAVOR, lpAction->ulActionFlavor, szFlags);
	*PropString += szTmp;

	delete[] szFlags;
	szFlags = NULL;

	if (!lpAction->lpPropTagArray)
	{
		szTmp.FormatMessage(IDS_ACTIONTAGARRAYNULL);
		*PropString += szTmp;
	}
	else
	{
		szTmp.FormatMessage(IDS_ACTIONTAGARRAYCOUNT, lpAction->lpPropTagArray->cValues);
		*PropString += szTmp;
		ULONG i = 0;
		for (i = 0; i < lpAction->lpPropTagArray->cValues; i++)
		{
			szTmp.FormatMessage(IDS_ACTIONTAGARRAYTAG,
				i,
				TagToString(lpAction->lpPropTagArray->aulPropTag[i], NULL, false, false));
			*PropString += szTmp;
		}
	}
} // ActionToString

void ActionsToString(_In_ ACTIONS* lpActions, _In_ CString* PropString)
{
	if (!PropString) return;

	*PropString = _T(""); // STRING_OK
	if (!lpActions)
	{
		PropString->FormatMessage(IDS_ACTIONSNULL);
		return;
	}
	CString szTmp;
	CString szAltTmp;

	LPTSTR szFlags = NULL;
	InterpretFlags(flagRulesVersion, lpActions->ulVersion, &szFlags);
	PropString->FormatMessage(IDS_ACTIONSMEMBERS,
		lpActions->ulVersion,
		szFlags,
		lpActions->cActions);
	delete[] szFlags;
	szFlags = NULL;

	UINT i = 0;
	for (i = 0; i < lpActions->cActions; i++)
	{
		szTmp.FormatMessage(IDS_ACTIONSACTION, i);
		*PropString += szTmp;
		ActionToString(&lpActions->lpAction[i], &szTmp);
		*PropString += szTmp;
	}
} // ActionsToString

void FileTimeToString(_In_ FILETIME* lpFileTime, _In_ CString *PropString, _In_opt_ CString *AltPropString)
{
	HRESULT	hRes = S_OK;
	SYSTEMTIME SysTime = { 0 };

	if (!lpFileTime) return;

	WC_B(FileTimeToSystemTime((FILETIME*)lpFileTime, &SysTime));

	if (S_OK == hRes && PropString)
	{
		int		iRet = 0;
		TCHAR	szTimeStr[MAX_PATH];
		TCHAR	szDateStr[MAX_PATH];
		CString szFormatStr;

		szTimeStr[0] = NULL;
		szDateStr[0] = NULL;

		// shove millisecond info into our format string since GetTimeFormat doesn't use it
		szFormatStr.FormatMessage(IDS_FILETIMEFORMAT, SysTime.wMilliseconds);

		WC_D(iRet, GetTimeFormat(
			LOCALE_USER_DEFAULT,
			NULL,
			&SysTime,
			szFormatStr,
			szTimeStr,
			MAX_PATH));

		WC_D(iRet, GetDateFormat(
			LOCALE_USER_DEFAULT,
			NULL,
			&SysTime,
			NULL,
			szDateStr,
			MAX_PATH));

		PropString->Format(_T("%s %s"), szTimeStr, szDateStr); // STRING_OK
	}
	else if (PropString)
	{
		PropString->FormatMessage(IDS_INVALIDSYSTIME);
	}

	if (AltPropString)
		AltPropString->FormatMessage(
		IDS_FILETIMEALTFORMAT,
		lpFileTime->dwLowDateTime,
		lpFileTime->dwHighDateTime);
} // FileTimeToString

/***************************************************************************
Name: InterpretProp
Purpose: Evaluate a property value and return a string representing the property.
Parameters:
In:
LPSPropValue lpProp: Property to be evaluated
Out:
wstring* tmpPropString: String representing property value
wstring* tmpAltPropString: Alternative string representation
Comment: Add new Property IDs as they become known
***************************************************************************/
void InterpretProp(_In_ LPSPropValue lpProp, _In_opt_  wstring* PropString, _In_opt_  wstring* AltPropString)
{
	if (!lpProp) return;

	Property parsedProperty = ParseProperty(lpProp);

	if (PropString) *PropString = parsedProperty.toString();
	if (AltPropString) *AltPropString = parsedProperty.toAltString();
}

_Check_return_ CString TypeToString(ULONG ulPropTag)
{
	CString tmpPropType;

	bool bNeedInstance = false;
	if (ulPropTag & MV_INSTANCE)
	{
		ulPropTag &= ~MV_INSTANCE;
		bNeedInstance = true;
	}

	ULONG ulCur = 0;
	bool bTypeFound = false;

	for (ulCur = 0; ulCur < ulPropTypeArray; ulCur++)
	{
		if (PropTypeArray[ulCur].ulValue == PROP_TYPE(ulPropTag))
		{
			tmpPropType = PropTypeArray[ulCur].lpszName;
			bTypeFound = true;
			break;
		}
	}
	if (!bTypeFound)
		tmpPropType.Format(_T("0x%04x"), PROP_TYPE(ulPropTag)); // STRING_OK

	if (bNeedInstance) tmpPropType += _T(" | MV_INSTANCE"); // STRING_OK
	return tmpPropType;
} // TypeToString

// TagToString will prepend the http://schemas.microsoft.com/MAPI/ for us since it's a constant
// We don't compute a DASL string for non-named props as FormatMessage in TagToString can handle those
void NameIDToStrings(_In_ LPMAPINAMEID lpNameID,
	ULONG ulPropTag,
	_In_ wstring& szPropName,
	_In_ wstring& szPropGUID,
	_In_ wstring& szDASL)
{
	HRESULT hRes = S_OK;

	// Can't generate strings without a MAPINAMEID structure
	if (!lpNameID) return;

	LPNAMEDPROPCACHEENTRY lpNamedPropCacheEntry = NULL;

	// If we're using the cache, look up the answer there and return
	if (fCacheNamedProps())
	{
		lpNamedPropCacheEntry = FindCacheEntry(PROP_ID(ulPropTag), lpNameID->lpguid, lpNameID->ulKind, lpNameID->Kind.lID, lpNameID->Kind.lpwstrName);
		if (lpNamedPropCacheEntry && lpNamedPropCacheEntry->bStringsCached)
		{
			szPropName = lpNamedPropCacheEntry->lpszPropName;
			szPropGUID = lpNamedPropCacheEntry->lpszPropGUID;
			szDASL = lpNamedPropCacheEntry->lpszDASL;
			return;
		}

		// We shouldn't ever get here without a cached entry
		if (!lpNamedPropCacheEntry)
		{
			DebugPrint(DBGNamedProp, _T("NameIDToStrings: Failed to find cache entry for ulPropTag = 0x%08X\n"), ulPropTag);
			return;
		}
	}

	DebugPrint(DBGNamedProp, _T("Parsing named property\n"));
	DebugPrint(DBGNamedProp, _T("ulPropTag = 0x%08x\n"), ulPropTag);
	szPropGUID = GUIDToWstringAndName(lpNameID->lpguid);
	DebugPrint(DBGNamedProp, _T("lpNameID->lpguid = %ws\n"), szPropGUID.c_str());

	wstring szDASLGuid = GUIDToWstring(lpNameID->lpguid);

	if (lpNameID->ulKind == MNID_ID)
	{
		DebugPrint(DBGNamedProp, _T("lpNameID->Kind.lID = 0x%04X = %d\n"), lpNameID->Kind.lID, lpNameID->Kind.lID);
		std::wstring szName = NameIDToPropName(lpNameID);

		if (!szName.empty())
		{
			// Printing hex first gets a nice sort without spacing tricks
			szPropName = format(L"id: 0x%04X=%d = %ws", // STRING_OK
				lpNameID->Kind.lID,
				lpNameID->Kind.lID,
				szName.c_str());

		}
		else
		{
			// Printing hex first gets a nice sort without spacing tricks
			szPropName = format(L"id: 0x%04X=%d", // STRING_OK
				lpNameID->Kind.lID,
				lpNameID->Kind.lID);
		}

		szDASL = format(L"id/%s/%04X%04X", // STRING_OK
			szDASLGuid.c_str(),
			lpNameID->Kind.lID,
			PROP_TYPE(ulPropTag));
	}
	else if (lpNameID->ulKind == MNID_STRING)
	{
		// lpwstrName is LPWSTR which means it's ALWAYS unicode
		// But some folks get it wrong and stuff ANSI data in there
		// So we check the string length both ways to make our best guess
		size_t cchShortLen = NULL;
		size_t cchWideLen = NULL;
		WC_H(StringCchLengthA((LPSTR)lpNameID->Kind.lpwstrName, STRSAFE_MAX_CCH, &cchShortLen));
		WC_H(StringCchLengthW(lpNameID->Kind.lpwstrName, STRSAFE_MAX_CCH, &cchWideLen));

		if (cchShortLen < cchWideLen)
		{
			// this is the *proper* case
			DebugPrint(DBGNamedProp, _T("lpNameID->Kind.lpwstrName = \"%ws\"\n"), lpNameID->Kind.lpwstrName);
			szPropName = format(L"sz: \"%ws\"", lpNameID->Kind.lpwstrName);

			szDASL = format(L"string/%ws/%ws", // STRING_OK
				szDASLGuid.c_str(),
				lpNameID->Kind.lpwstrName);
		}
		else
		{
			// this is the case where ANSI data was shoved into a unicode string.
			DebugPrint(DBGNamedProp, _T("Warning: ANSI data was found in a unicode field. This is a bug on the part of the creator of this named property\n"));
			DebugPrint(DBGNamedProp, _T("lpNameID->Kind.lpwstrName = \"%hs\"\n"), (LPCSTR)lpNameID->Kind.lpwstrName);

			wstring szComment = loadstring(IDS_NAMEWASANSI);
			szPropName = format(L"sz: \"%hs\" %ws", (LPSTR)lpNameID->Kind.lpwstrName, szComment);

			szDASL = format(L"string/%ws/%hs", // STRING_OK
				szDASLGuid.c_str(),
				lpNameID->Kind.lpwstrName);
		}
	}

	// We've built our strings - if we're caching, put them in the cache
	if (lpNamedPropCacheEntry)
	{
		lpNamedPropCacheEntry->lpszPropName = szPropName;
		lpNamedPropCacheEntry->lpszPropGUID = szPropGUID;
		lpNamedPropCacheEntry->lpszDASL = szDASL;
		lpNamedPropCacheEntry->bStringsCached = true;
	}
}

// lpszNamedPropName, lpszNamedPropGUID, lpszNamedPropDASL freed with FreeNameIDStrings
void NameIDToStrings(
	ULONG ulPropTag, // optional 'original' prop tag
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	_Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
	_Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
	_Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropDASL) // Built from ulPropTag & lpMAPIProp
{
	HRESULT hRes = S_OK;

	// In case we error out, set our returns
	if (lpszNamedPropName) *lpszNamedPropName = NULL;
	if (lpszNamedPropGUID) *lpszNamedPropGUID = NULL;
	if (lpszNamedPropDASL) *lpszNamedPropDASL = NULL;

	// Named Props
	LPMAPINAMEID* lppPropNames = 0;

	// If we weren't passed named property information and we need it, look it up
	// We check bIsAB here - some address book providers return garbage which will crash us
	if (!lpNameID &&
		lpMAPIProp && // if we have an object
		!bIsAB &&
		RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD && // and we're parsing named props
		(RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD || PROP_ID(ulPropTag) >= 0x8000) && // and it's either a named prop or we're doing all props
		(lpszNamedPropName || lpszNamedPropGUID || lpszNamedPropDASL)) // and we want to return something that needs named prop information
	{
		SPropTagArray tag = { 0 };
		LPSPropTagArray lpTag = &tag;
		ULONG ulPropNames = 0;
		tag.cValues = 1;
		tag.aulPropTag[0] = ulPropTag;

		WC_H_GETPROPS(GetNamesFromIDs(lpMAPIProp,
			lpMappingSignature,
			&lpTag,
			NULL,
			NULL,
			&ulPropNames,
			&lppPropNames));
		if (SUCCEEDED(hRes) && ulPropNames == 1 && lppPropNames && lppPropNames[0])
		{
			lpNameID = lppPropNames[0];
		}
		hRes = S_OK;
	}

	if (lpNameID)
	{
		wstring lpszPropName;
		wstring lpszPropGUID;
		wstring lpszDASL;

		NameIDToStrings(lpNameID,
			ulPropTag,
			lpszPropName,
			lpszPropGUID,
			lpszDASL);

		if (lpszNamedPropName) *lpszNamedPropName = wstringToLPTSTR(lpszPropName);
		if (lpszNamedPropGUID) *lpszNamedPropGUID = wstringToLPTSTR(lpszPropGUID);
		if (lpszNamedPropDASL) *lpszNamedPropDASL = wstringToLPTSTR(lpszDASL);
	}

	// Avoid making the call if we don't have to so we don't accidently depend on MAPI
	if (lppPropNames) MAPIFreeBuffer(lppPropNames);
}

// Free strings from NameIDToStrings if necessary
// If we're using the cache, we don't need to free
// Need to watch out for callers to NameIDToStrings holding the strings
// long enough for the user to change the cache setting!
void FreeNameIDStrings(_In_opt_z_ LPTSTR lpszPropName,
	_In_opt_z_ LPTSTR lpszPropGUID,
	_In_opt_z_ LPTSTR lpszDASL)
{
	delete[] lpszPropName;
	delete[] lpszPropGUID;
	delete[] lpszDASL;
}