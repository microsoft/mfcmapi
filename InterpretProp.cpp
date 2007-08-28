#include "stdafx.h"
#include "Error.h"

#include "InterpretProp.h"

#include "MAPIFunctions.h"
#include "Registry.h"
#include "InterpretProp2.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

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

//allocates output buffer with new
//delete with delete[]
//suprisingly, this algorithm works in a unicode build as well
HRESULT Base64Decode(LPCTSTR szEncodedStr, size_t* cbBuf, LPBYTE* lpDecodedBuffer)
{
	HRESULT hRes = S_OK;
	size_t	cchLen = 0;

	EC_H(StringCchLength(szEncodedStr,STRSAFE_MAX_CCH,&cchLen));

	if (cchLen % 4) return MAPI_E_INVALID_PARAMETER;

	//look for padding at the end
	static const TCHAR szPadding[]  = _T("==");// STRING_OK
	const TCHAR* szPaddingLoc = NULL;
	szPaddingLoc = _tcschr(szEncodedStr, szPadding[0]);
	size_t cchPaddingLen = 0;
	if (NULL != szPaddingLoc)
	{
		//check padding length
		EC_H(StringCchLength(szPaddingLoc,STRSAFE_MAX_CCH,&cchPaddingLen));
		if (cchPaddingLen >= 3) return MAPI_E_INVALID_PARAMETER;

		//check for bad characters after the first '='
		if (_tcsncmp(szPaddingLoc, (TCHAR *) szPadding, cchPaddingLen)) return MAPI_E_INVALID_PARAMETER;
	}
	//cchPaddingLen == 0,1,2 now

	size_t	cchDecodedLen = ((cchLen + 3)/ 4) * 3;//3 times number of 4 tuplets, rounded up

	//back off the decoded length to the correct length
	// xx== ->y
	// xxx= ->yY
	// x=== ->this is a bad case which should never happen
	cchDecodedLen -= cchPaddingLen;
	//we have no room for error now!
	*lpDecodedBuffer = new BYTE[cchDecodedLen];
	if (!*lpDecodedBuffer) return MAPI_E_CALL_FAILED;

	*cbBuf = cchDecodedLen;

	LPBYTE	lpOutByte = *lpDecodedBuffer;

	TCHAR c[4] = {0};
	BYTE bTmp[3] = {0};//output

	while (*szEncodedStr)
	{
		int i = 0;
		int iOutlen = 3;
		for (i = 0 ; i < 4 ; i++)
		{
			c[i] = *(szEncodedStr+i);
			if  (c[i] == _T('='))
			{
				iOutlen = i-1;
				break;
			}
			if ((c[i] < 0x2b) || (c[i] > 0x7a)) return MAPI_E_INVALID_PARAMETER;

			c[i] = pBase64[c[i] - 0x2b];
		}
		bTmp[0]  = (BYTE) ((c[0] << 2)        | (c[1] >> 4));
		bTmp[1]  = (BYTE) ((c[1] & 0x0f) << 4 | (c[2] >> 2));
		bTmp[2]  = (BYTE) ((c[2] & 0x03) << 6 |  c[3]);

		for (i = 0 ; i < iOutlen ; i++)
		{
			lpOutByte[i] = bTmp[i];
		}
		lpOutByte += 3;
		szEncodedStr += 4;
	}

	return hRes;
}

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

//allocates output string with new
//delete with delete[]
HRESULT Base64Encode(size_t cbSourceBuf, LPBYTE lpSourceBuffer, size_t* cchEncodedStr, LPTSTR* szEncodedStr)
{
	HRESULT hRes = S_OK;

	size_t cchEncodeLen = ((cbSourceBuf + 2) / 3) * 4;//4 * number of size three blocks, round up, plus null terminator
	*szEncodedStr = new TCHAR[cchEncodeLen+1];//allocate a touch extra for some NULL terminators
	if (cchEncodedStr) *cchEncodedStr = cchEncodeLen;
	if (!*szEncodedStr) return MAPI_E_CALL_FAILED;

	size_t cbBuf = 0; // General purpose integers.
	TCHAR* szOutChar = NULL;
	szOutChar = *szEncodedStr;

	// Using integer division to round down here
	while (cbBuf < (cbSourceBuf/3)*3) // encode each 3 byte octet.
	{
		*szOutChar++ = pIndex[  lpSourceBuffer[cbBuf]             >> 2];
		*szOutChar++ = pIndex[((lpSourceBuffer[cbBuf]     & 0x03) << 4) + (lpSourceBuffer[cbBuf + 1] >> 4)];
		*szOutChar++ = pIndex[((lpSourceBuffer[cbBuf + 1] & 0x0f) << 2) + (lpSourceBuffer[cbBuf + 2] >> 6)];
		*szOutChar++ = pIndex[  lpSourceBuffer[cbBuf + 2] & 0x3f];
		cbBuf    += 3; // Next octet.
	}

	if (cbSourceBuf - cbBuf) // Partial octet remaining?
	{
		*szOutChar++ = pIndex[lpSourceBuffer[cbBuf] >> 2]; // Yes, encode it.

		if  (cbSourceBuf - cbBuf == 1) // End of octet?
		{
			*szOutChar++ = pIndex[ (lpSourceBuffer[cbBuf] & 0x03) << 4];
			*szOutChar++ = _T('=');
			*szOutChar++ = _T('=');
		}
		else
		{ // No, one more part.
			*szOutChar++ = pIndex[((lpSourceBuffer[cbBuf]     & 0x03) << 4) + (lpSourceBuffer[cbBuf + 1] >> 4)];
			*szOutChar++ = pIndex[ (lpSourceBuffer[cbBuf + 1] & 0x0f) << 2];
			*szOutChar++ = _T('=');
		}
	}
	*szOutChar = _T('\0');

	return hRes;
}

CString BinToTextString(LPSBinary lpBin, BOOL bMultiLine)
{
	if (!lpBin) return _T("");

	CString StringAsText;
	LPTSTR szBin = NULL;

	szBin = new TCHAR[1+lpBin->cb];

	if (szBin)
	{
		ULONG i;
		for (i = 0;i<lpBin->cb;i++)
		{
			if (lpBin->lpb[i] >= 0x20 &&
				lpBin->lpb[i] <= 0xFE)
			{
				szBin[i] = lpBin->lpb[i];
			}
			else if (bMultiLine &&
				(lpBin->lpb[i] == 8 ||
				lpBin->lpb[i] == 9 ||
				lpBin->lpb[i] == 10 ||
				lpBin->lpb[i] == 13))
			{
				szBin[i] = lpBin->lpb[i];
			}
			else
			{
				szBin[i] = _T('.');
			}

		}
		szBin[i] = _T('\0');

		StringAsText = szBin;

		delete[] szBin;
	}
	return StringAsText;
}

CString BinToHexString(LPSBinary lpBin, BOOL bPrependCB)
{
	if (!lpBin) return _T("");

	CString HexString;

	if (!lpBin->cb)
	{
		if (bPrependCB)
		{
			HexString.Format(_T("cb: 0 lpb: NULL"));// STRING_OK
		}
		return HexString;
	}

	LPTSTR szBin = NULL;
	MyHexFromBin(
		lpBin->lpb,
		lpBin->cb,
		&szBin);

	if (szBin)
	{
		if (bPrependCB)
		{
			HexString.Format(_T("cb: %d lpb: %s"),lpBin->cb,szBin);// STRING_OK
		}
		else
		{
			HexString = szBin;
		}
		delete[] szBin;
	}
	return HexString;
}

//Allocates string for GUID with new
//free with delete[]
#define GUID_STRING_SIZE 39
LPTSTR GUIDToString(LPCGUID lpGUID)
{
	HRESULT	hRes = S_OK;
	GUID	nullGUID = {0};
	LPTSTR	szGUID = NULL;

	if (!lpGUID)
	{
		lpGUID = &nullGUID;
	}

	szGUID = new TCHAR[GUID_STRING_SIZE];

	EC_H(StringCchPrintf(szGUID,GUID_STRING_SIZE,_T("{%.8X-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X}"),// STRING_OK
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
}

HRESULT StringToGUID(LPCTSTR szGUID, LPGUID lpGUID)
{
	HRESULT hRes = S_OK;
	if (!szGUID || !lpGUID) return MAPI_E_INVALID_PARAMETER;

    WCHAR szWGuid[GUID_STRING_SIZE];
    LPCTSTR pszSrc = NULL;
    LPWSTR pszWDest = NULL;

    pszSrc = szGUID;

    if (*pszSrc != '{') return MAPI_E_INVALID_PARAMETER;

    // Convert to Unicode while you are copying to your temporary buffer.
    // Do not worry about non-ANSI characters; this is a GUID string.
    pszWDest = szWGuid;

    while((*pszSrc) && (*pszSrc != '}') &&
		(pszWDest < &szWGuid[GUID_STRING_SIZE - 2]))
	{
        *pszWDest++ = *pszSrc++;
    }

    // On success, pszSrc will point to '}' (the last character of the GUID string).
    if (*pszSrc != '}') return MAPI_E_INVALID_PARAMETER;

    // pszDest will still be in range and have two chars left because
    // of the condition in the preceding while loop.
    *pszWDest++ = '}';
    *pszWDest = '\0';

    // Borrow the functionality of CLSIDFromString to get the 16-byte
    // GUID from the GUID string.
	WC_H(CLSIDFromString(
        szWGuid,
        lpGUID));

    return hRes;
}

CString CurrencyToString(CURRENCY curVal)
{
	CString szCur;

	szCur.Format(_T("%05I64d"),curVal.int64);// STRING_OK
	if (szCur.GetLength() > 4)
	{
		szCur.Insert(szCur.GetLength()-4,_T("."));// STRING_OK
	}
	return szCur;
}

CString TagToString(ULONG ulPropTag, LPMAPIPROP lpObj, BOOL bIsAB, BOOL bSingleLine)
{
	HRESULT hRes = S_OK;
	CString szRet;
	CString szTemp;

	LPTSTR szName = NULL;
	LPTSTR szGuid = NULL;
	LPTSTR szDASL = NULL;
	LPTSTR szExactMatches = NULL;
	LPTSTR szPartialMatches = NULL;
	EC_H(PropTagToPropName(ulPropTag,bIsAB,&szExactMatches,&szPartialMatches));

	GetPropName(lpObj,ulPropTag,&szName,&szGuid,&szDASL);

	CString szFormatString;
	if (bSingleLine)
	{
		szFormatString = _T("0x%1!08X! (%2)");// STRING_OK
		if (szExactMatches) szFormatString += _T(": %3");// STRING_OK
		if (szPartialMatches) szFormatString += _T(": (%4)");// STRING_OK
		if (szName)
		{
			szTemp.LoadString(IDS_NAMEDPROPSINGLELINE);
			szFormatString += szTemp;
		}
		if (szGuid)
		{
			szTemp.LoadString(IDS_GUIDSINGLELINE);
			szFormatString += szTemp;
		}
	}
	else
	{
		szFormatString.LoadString(IDS_TAGMULTILINE);
		if (szExactMatches)
		{
			szTemp.LoadString(IDS_PROPNAMEMULTILINE);
			szFormatString += szTemp;
		}
		if (szPartialMatches)
		{
			szTemp.LoadString(IDS_OTHERNAMESMULTILINE);
			szFormatString += szTemp;
		}
		if (PROP_ID(ulPropTag) < 0x8000)
		{
			szTemp.LoadString(IDS_DASLPROPTAG);
			szFormatString += szTemp;
		}
		else
		{
			szTemp.LoadString(IDS_DASLNAMED);
			szFormatString += szTemp;
		}
		if (szName)
		{
			szTemp.LoadString(IDS_NAMEPROPNAMEMULTILINE);
			szFormatString += szTemp;
		}
		if (szGuid)
		{
			szTemp.LoadString(IDS_NAMEPROPGUIDMULTILINE);
			szFormatString += szTemp;
		}
	}
	szRet.FormatMessage(szFormatString,
		ulPropTag,
		(LPCTSTR) TypeToString(ulPropTag),
		szExactMatches,
		szPartialMatches,
		szName,
		szGuid,
		szDASL);

	delete[] szPartialMatches;
	delete[] szExactMatches;
	delete[] szDASL;
	delete[] szGuid;
	delete[] szName;

	if (fIsSet(DBGTest))
	{
		static size_t cchMaxBuff = 0;
		size_t cchBuff = szRet.GetLength();
		cchMaxBuff = max(cchBuff,cchMaxBuff);
		DebugPrint(DBGTest,_T("TagToString parsing 0x%08X returned %d chars - max %d\n"),ulPropTag,cchBuff,cchMaxBuff);
	}
	return szRet;
}

CString ProblemArrayToString(LPSPropProblemArray lpProblems)
{
	CString szOut;
	if (lpProblems)
	{
		ULONG i = 0;
		for (i = 0;i < lpProblems->cProblem;i++)
		{
			CString szTemp;
			szTemp.FormatMessage(IDS_PROBLEMARRAY,
				lpProblems->aProblem[i].ulIndex,
				TagToString(lpProblems->aProblem[i].ulPropTag,NULL,false,false),
				lpProblems->aProblem[i].scode,
				ErrorNameFromErrorCode(lpProblems->aProblem[i].scode));
			szOut += szTemp;
		}
	}
	return szOut;
}

CString MAPIErrToString(ULONG ulFlags, LPMAPIERROR lpErr)
{
	CString szOut;
	if (lpErr)
	{
		szOut.FormatMessage(
			ulFlags & MAPI_UNICODE?IDS_MAPIERRUNICODE:IDS_MAPIERRANSI,
			lpErr->ulVersion,
			lpErr->lpszError,
			lpErr->lpszComponent,
			lpErr->ulLowLevelError,
			ErrorNameFromErrorCode(lpErr->ulLowLevelError),
			lpErr->ulContext);
	}
	return szOut;
}

CString TnefProblemArrayToString(LPSTnefProblemArray lpError)
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
				TagToString(lpError->aProblem[iError].ulPropTag,NULL,false,false),
				lpError->aProblem[iError].scode,
				ErrorNameFromErrorCode(lpError->aProblem[iError].scode));
			szOut += szTemp;
		}
	}
	return szOut;
}

CString EntryListToString(LPENTRYLIST lpEntryList)
{
	CString szRet;
	if (lpEntryList)
	{
		szRet.FormatMessage(IDS_ENTRYLISTPREFIX,lpEntryList->cValues);
		ULONG i = 0;

		for (i = 0; i < lpEntryList->cValues; i++)
		{
			CString szTemp;
			szTemp.FormatMessage(
				IDS_ENTRYLISTENTRY,
				BinToHexString(&lpEntryList->lpbin[i],true),
				BinToTextString(&lpEntryList->lpbin[i],false));
			szRet+= szTemp;
		}
	}
	else szRet.FormatMessage(IDS_NULLENTRYLIST);
	return szRet;
}

void RestrictionToString(LPSRestriction lpRes, LPMAPIPROP lpObj, ULONG ulTabLevel, CString *PropString)
{
	if (!PropString) return;

	*PropString = _T("");// STRING_OK

	ULONG i = 0;
	if (!lpRes)
	{
		PropString->FormatMessage(IDS_NULLRES);
		return;
	}
	CString szTmp;
	CString szProp;
	CString szAltProp;

	CString szTabs;
	for (i = 0;i<ulTabLevel;i++)
	{
		szTabs += _T("\t");// STRING_OK
	}

	HRESULT hRes = S_OK;
	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagRestrictionType, lpRes->rt, &szFlags));
	szTmp.FormatMessage(IDS_RESTYPE,szTabs,szFlags);
	*PropString += szTmp;
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;
	switch(lpRes->rt)
	{
	case RES_COMPAREPROPS:
		EC_H(InterpretFlags(flagRelop, lpRes->res.resCompareProps.relop, &szFlags));
		szTmp.FormatMessage(
			IDS_RESCOMPARE,
			szTabs,
			szFlags,
			lpRes->res.resCompareProps.relop,
			TagToString(lpRes->res.resCompareProps.ulPropTag1,lpObj,false,true),
			TagToString(lpRes->res.resCompareProps.ulPropTag2,lpObj,false,true));
		*PropString += szTmp;
		MAPIFreeBuffer(szFlags);
		break;
	case RES_AND:
		szTmp.FormatMessage(IDS_RESANDCOUNT,szTabs,lpRes->res.resAnd.cRes);
		*PropString += szTmp;
		for (i = 0;i< lpRes->res.resAnd.cRes;i++)
		{
			szTmp.FormatMessage(IDS_RESANDPOINTER,szTabs,i,&lpRes->res.resAnd.lpRes[i]);
			*PropString += szTmp;
			RestrictionToString(&lpRes->res.resAnd.lpRes[i],lpObj,ulTabLevel+1,&szTmp);
			*PropString += szTmp;
		}
		break;
	case RES_OR:
		szTmp.FormatMessage(IDS_RESORCOUNT,szTabs,lpRes->res.resOr.cRes);
		*PropString += szTmp;
		for (i = 0;i< lpRes->res.resOr.cRes;i++)
		{
			szTmp.FormatMessage(IDS_RESORPOINTER,szTabs,i,&lpRes->res.resOr.lpRes[i]);
			*PropString += szTmp;
			RestrictionToString(&lpRes->res.resOr.lpRes[i],lpObj,ulTabLevel+1,&szTmp);
			*PropString += szTmp;
		}
		break;
	case RES_NOT:
		szTmp.FormatMessage(
			IDS_RESNOT,
			szTabs,
			lpRes->res.resNot.ulReserved,
			lpRes->res.resNot.lpRes);
		*PropString += szTmp;
		RestrictionToString(lpRes->res.resNot.lpRes,lpObj,ulTabLevel+1,&szTmp);
		*PropString += szTmp;
		break;
	case RES_CONTENT:
		EC_H(InterpretFlags(flagFuzzyLevel, lpRes->res.resContent.ulFuzzyLevel, &szFlags));
		szTmp.FormatMessage(
			IDS_RESCONTENT,
			szTabs,
			szFlags,
			lpRes->res.resContent.ulFuzzyLevel,
			TagToString(lpRes->res.resContent.ulPropTag,lpObj,false,true));
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
		*PropString += szTmp;
		if (lpRes->res.resContent.lpProp)
		{
			InterpretProp(lpRes->res.resContent.lpProp,&szProp,&szAltProp);
			szTmp.FormatMessage(
				IDS_RESCONTENTPROP,
				szTabs,
				TagToString(lpRes->res.resContent.lpProp->ulPropTag,lpObj,false,true),
				szProp,
				szAltProp);
			*PropString += szTmp;
		}
		break;
	case RES_PROPERTY:
		EC_H(InterpretFlags(flagRelop, lpRes->res.resProperty.relop, &szFlags));
		szTmp.FormatMessage(
			IDS_RESPROP,
			szTabs,
			szFlags,
			lpRes->res.resProperty.relop,
			TagToString(lpRes->res.resProperty.ulPropTag,lpObj,false,true));
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
		*PropString += szTmp;
		if (lpRes->res.resProperty.lpProp)
		{
			InterpretProp(lpRes->res.resProperty.lpProp,&szProp,&szAltProp);
			szTmp.FormatMessage(
				IDS_RESPROPPROP,
				szTabs,
				TagToString(lpRes->res.resProperty.lpProp->ulPropTag,lpObj,false,true),
				szProp,
				szAltProp);
			*PropString += szTmp;
			EC_H(InterpretFlags(lpRes->res.resProperty.lpProp, &szFlags));
			if (szFlags)
			{
				szTmp.FormatMessage(IDS_RESPROPPROPFLAGS,szTabs,szFlags);
				MAPIFreeBuffer(szFlags);
				szFlags = NULL;
				*PropString += szTmp;
			}
		}
		break;
	case RES_BITMASK:
		EC_H(InterpretFlags(flagBitmask, lpRes->res.resBitMask.relBMR, &szFlags));
		szTmp.FormatMessage(
			IDS_RESBITMASK,
			szTabs,
			szFlags,
			lpRes->res.resBitMask.relBMR,
			lpRes->res.resBitMask.ulMask);
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
		*PropString += szTmp;
		EC_H(InterpretFlags(PROP_ID(lpRes->res.resBitMask.ulPropTag), lpRes->res.resBitMask.ulMask, &szFlags));
		if (szFlags)
		{
			szTmp.FormatMessage(IDS_RESBITMASKFLAGS,szFlags);
			MAPIFreeBuffer(szFlags);
			szFlags = NULL;
			*PropString += szTmp;
		}
		szTmp.FormatMessage(
			IDS_RESBITMASKTAG,
			szTabs,
			TagToString(lpRes->res.resBitMask.ulPropTag,lpObj,false,true));
		*PropString += szTmp;
		break;
	case RES_SIZE:
		EC_H(InterpretFlags(flagRelop, lpRes->res.resSize.relop, &szFlags));
		szTmp.FormatMessage(
			IDS_RESSIZE,
			szTabs,
			szFlags,
			lpRes->res.resSize.relop,
			lpRes->res.resSize.cb,
			TagToString(lpRes->res.resSize.ulPropTag,lpObj,false,true));
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
		*PropString += szTmp;
		break;
	case RES_EXIST:
		szTmp.FormatMessage(
			IDS_RESEXIST,
			szTabs,
			TagToString(lpRes->res.resExist.ulPropTag,lpObj,false,true),
			lpRes->res.resExist.ulReserved1,
			lpRes->res.resExist.ulReserved2);
		*PropString += szTmp;
		break;
	case RES_SUBRESTRICTION:
		szTmp.FormatMessage(
			IDS_RESSUBRES,
			szTabs,
			TagToString(lpRes->res.resSub.ulSubObject,lpObj,false,true),
			lpRes->res.resSub.lpRes);
		*PropString += szTmp;
		RestrictionToString(lpRes->res.resSub.lpRes,lpObj,ulTabLevel+1,&szTmp);
		*PropString += szTmp;
		break;
	case RES_COMMENT:
		szTmp.FormatMessage(IDS_RESCOMMENT,szTabs,lpRes->res.resComment.cValues);
		*PropString += szTmp;
		for (i = 0;i< lpRes->res.resComment.cValues;i++)
		{
			InterpretProp(&lpRes->res.resComment.lpProp[i],&szProp,&szAltProp);
			szTmp.FormatMessage(
				IDS_RESCOMMENTPROPS,
				szTabs,
				i,
				TagToString(lpRes->res.resComment.lpProp[i].ulPropTag,lpObj,false,true),
				szProp,
				szAltProp);
			*PropString += szTmp;
		}
		szTmp.FormatMessage(
			IDS_RESCOMMENTRES,
			szTabs,
			lpRes->res.resComment.lpRes);
		*PropString += szTmp;
		RestrictionToString(lpRes->res.resComment.lpRes,lpObj,ulTabLevel+1,&szTmp);
		*PropString += szTmp;
		break;
	}
}

CString RestrictionToString(LPSRestriction lpRes, LPMAPIPROP lpObj)
{
	CString szRes;
	RestrictionToString(lpRes,lpObj,0,&szRes);
	return szRes;
}

void AdrListToString(LPADRLIST lpAdrList,CString *PropString)
{
	if (!PropString) return;

	*PropString = _T("");// STRING_OK
	if (!lpAdrList)
	{
		PropString->FormatMessage(IDS_ADRLISTNULL);
		return;
	}
	CString szTmp;
	CString szProp;
	CString szAltProp;
	PropString->FormatMessage(IDS_ADRLISTCOUNT,lpAdrList->cEntries);

	ULONG i = 0;
	for (i = 0 ; i < lpAdrList->cEntries ; i++)
	{
		szTmp.FormatMessage(IDS_ADRLISTENTRIESCOUNT,i,lpAdrList->aEntries[i].cValues);
		*PropString += szTmp;

		ULONG j = 0;
		for (j = 0 ; j < lpAdrList->aEntries[i].cValues ; j++)
		{
			InterpretProp(&lpAdrList->aEntries[i].rgPropVals[j],&szProp,&szAltProp);
			szTmp.FormatMessage(
				IDS_ADRLISTENTRY,
				i,
				j,
				TagToString(lpAdrList->aEntries[i].rgPropVals[j].ulPropTag,NULL,false,false),
				szProp,
				szAltProp);
			*PropString += szTmp;
		}
	}
}

void ActionToString(ACTION* lpAction, CString* PropString)
{
	if (!PropString) return;

	*PropString = _T("");// STRING_OK
	if (!lpAction)
	{
		PropString->FormatMessage(IDS_ACTIONNULL);
		return;
	}
	CString szTmp;
	CString szProp;
	CString szAltProp;
	HRESULT hRes = S_OK;
	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagAccountType, lpAction->acttype, &szFlags));
	PropString->FormatMessage(
		IDS_ACTION,
		lpAction->acttype,
		szFlags,
		RestrictionToString(lpAction->lpRes,NULL),
		lpAction->ulFlags);
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;

	switch(lpAction->acttype)
	{
	case OP_MOVE:
	case OP_COPY:
		{
			SBinary sBinStore = {0};
			SBinary sBinFld = {0};
			sBinStore.cb = lpAction->actMoveCopy.cbStoreEntryId;
			sBinStore.lpb = (LPBYTE) lpAction->actMoveCopy.lpStoreEntryId;
			sBinFld.cb = lpAction->actMoveCopy.cbFldEntryId;
			sBinFld.lpb = (LPBYTE) lpAction->actMoveCopy.lpFldEntryId;

			szTmp.FormatMessage(IDS_ACTIONOPMOVECOPY,
				BinToHexString(&sBinStore,true),
				BinToTextString(&sBinStore,false),
				BinToHexString(&sBinFld,true),
				BinToTextString(&sBinFld,false));
			*PropString += szTmp;
			break;
		}
	case OP_REPLY:
	case OP_OOF_REPLY:
		{

			SBinary sBin = {0};
			sBin.cb = lpAction->actReply.cbEntryId;
			sBin.lpb = (LPBYTE) lpAction->actReply.lpEntryId;
			LPTSTR szGUID = GUIDToStringAndName(&lpAction->actReply.guidReplyTemplate);

			szTmp.FormatMessage(IDS_ACTIONOPREPLY,
				BinToHexString(&sBin,true),
				BinToTextString(&sBin,false),
				szGUID);
			*PropString += szTmp;
			delete[] szGUID;
			break;
		}
	case OP_DEFER_ACTION:
		{
			SBinary sBin = {0};
			sBin.cb = lpAction->actDeferAction.cbData;
			sBin.lpb = (LPBYTE) lpAction->actDeferAction.pbData;

			szTmp.FormatMessage(IDS_ACTIONOPDEFER,
				BinToHexString(&sBin,true),
				BinToTextString(&sBin,false));
			*PropString += szTmp;
			break;
		}
	case OP_BOUNCE:
		{
			EC_H(InterpretFlags(flagBounceCode, lpAction->scBounceCode, &szFlags));
			szTmp.FormatMessage(IDS_ACTIONOPBOUNCE,lpAction->scBounceCode,szFlags);
			MAPIFreeBuffer(szFlags);
			szFlags = NULL;
			*PropString += szTmp;
			break;
		}
	case OP_FORWARD:
	case OP_DELEGATE:
		{
			AdrListToString(lpAction->lpadrlist,&szProp);
			szTmp.FormatMessage(IDS_ACTIONOPFORWARDDEL,szProp);
			*PropString += szTmp;
			break;
		}

	case OP_TAG:
		{
			InterpretProp(&lpAction->propTag,&szProp,&szAltProp);
			szTmp.FormatMessage(IDS_ACTIONOPTAG,
				TagToString(lpAction->propTag.ulPropTag,NULL,false,true),
				szProp,
				szAltProp);
			*PropString += szTmp;
			break;
		}
	}

	switch(lpAction->acttype)
	{
	case OP_REPLY:
		{
			EC_H(InterpretFlags(flagOPReply, lpAction->ulActionFlavor, &szFlags));
			break;
		}
	case OP_FORWARD:
		{
			EC_H(InterpretFlags(flagOpForward, lpAction->ulActionFlavor, &szFlags));
			break;
		}
	}
	szTmp.FormatMessage(IDS_ACTIONFLAVOR,lpAction->ulActionFlavor,szFlags);
	*PropString += szTmp;

	MAPIFreeBuffer(szFlags);
	szFlags = NULL;

	if (!lpAction->lpPropTagArray)
	{
		szTmp.FormatMessage(IDS_ACTIONTAGARRAYNULL);
		*PropString += szTmp;
	}
	else
	{
		szTmp.FormatMessage(IDS_ACTIONTAGARRAYCOUNT,lpAction->lpPropTagArray->cValues);
		*PropString += szTmp;
		ULONG i = 0;
		for (i = 0 ; i < lpAction->lpPropTagArray->cValues ; i++)
		{
			szTmp.FormatMessage(IDS_ACTIONTAGARRAYTAG,
				i,
				TagToString(lpAction->lpPropTagArray->aulPropTag[i],NULL,false,false));
			*PropString += szTmp;
		}
	}
}

void ActionsToString(ACTIONS* lpActions, CString* PropString)
{
	if (!PropString) return;

	*PropString = _T("");// STRING_OK
	if (!lpActions)
	{
		PropString->FormatMessage(IDS_ACTIONSNULL);
		return;
	}
	CString szTmp;
	CString szAltTmp;

	HRESULT hRes = S_OK;
	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagRulesVersion, lpActions->ulVersion, &szFlags));
	PropString->FormatMessage(IDS_ACTIONSMEMBERS,
		lpActions->ulVersion,
		szFlags,
		lpActions->cActions);
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;

	UINT i = 0;
	for (i = 0 ; i < lpActions->cActions ; i++)
	{
		szTmp.FormatMessage(IDS_ACTIONSACTION,i);
		*PropString += szTmp;
		ActionToString(&lpActions->lpAction[i],&szTmp);
		*PropString += szTmp;
	}
}


void FileTimeToString(FILETIME* lpFileTime,CString *PropString,CString *AltPropString)
{
	HRESULT	hRes = S_OK;
	SYSTEMTIME SysTime = {0};

	if (!lpFileTime) return;

	WC_B(FileTimeToSystemTime((FILETIME FAR *) lpFileTime,&SysTime));

	if (S_OK == hRes && PropString)
	{
		int		iRet = 0;
		TCHAR	szTimeStr[MAX_PATH];
		TCHAR	szDateStr[MAX_PATH];
		CString szFormatStr;

		szTimeStr[0] = NULL;
		szDateStr[0] = NULL;

		//shove millisecond info into our format string since GetTimeFormat doesn't use it
		szFormatStr.FormatMessage(IDS_FILETIMEFORMAT,SysTime.wMilliseconds);

		WC_D(iRet,GetTimeFormat(
			LOCALE_USER_DEFAULT,
			NULL,
			&SysTime,
			szFormatStr,
			szTimeStr,
			MAX_PATH));

		WC_D(iRet,GetDateFormat(
			LOCALE_USER_DEFAULT,
			NULL,
			&SysTime,
			NULL,
			szDateStr,
			MAX_PATH));

		PropString->Format(_T("%s %s"),szTimeStr,szDateStr);// STRING_OK
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
}

void InterpretMVProp(LPSPropValue lpProp, ULONG ulMVRow, CString *PropString,CString *AltPropString)
{
	CString szTmp;
	CString	szAltTmp;

	if (!lpProp) return;
	if (ulMVRow > lpProp->Value.MVi.cValues) return;

	switch(PROP_TYPE(lpProp->ulPropTag))
	{
	case(PT_MV_I2):
		szTmp.Format(_T("%d"),lpProp->Value.MVi.lpi[ulMVRow]);// STRING_OK
		szAltTmp.Format(_T("0x%X"),lpProp->Value.MVi.lpi[ulMVRow]);// STRING_OK
		break;
	case(PT_MV_LONG):
		szTmp.Format(_T("%u"),lpProp->Value.MVl.lpl[ulMVRow]);// STRING_OK
		szAltTmp.Format(_T("0x%X"),lpProp->Value.MVl.lpl[ulMVRow]);// STRING_OK
		break;
	case(PT_MV_DOUBLE):
		szTmp.Format(_T("%u"),lpProp->Value.MVdbl.lpdbl[ulMVRow]);// STRING_OK
		break;
	case(PT_MV_CURRENCY):
		szTmp = CurrencyToString(lpProp->Value.MVcur.lpcur[ulMVRow]);
		szAltTmp.Format(_T("0x%08X:0x%08X"),(int)(lpProp->Value.MVcur.lpcur[ulMVRow].Hi),(int)lpProp->Value.MVcur.lpcur[ulMVRow].Lo);// STRING_OK
		break;
	case(PT_MV_APPTIME):
		szTmp.Format(_T("%u"),lpProp->Value.MVat.lpat[ulMVRow]);// STRING_OK
		break;
	case(PT_MV_SYSTIME):
		FileTimeToString(&lpProp->Value.MVft.lpft[ulMVRow],&szTmp,&szAltTmp);
		break;
	case(PT_MV_I8):
		szTmp.Format(_T("%I64d"),lpProp->Value.MVli.lpli[ulMVRow].QuadPart);// STRING_OK
		szAltTmp.Format(_T("0x%08X:0x%08X"),(int)(lpProp->Value.MVli.lpli[ulMVRow].HighPart),(int)lpProp->Value.MVli.lpli[ulMVRow].LowPart);// STRING_OK
		break;
	case(PT_MV_R4):
		szTmp.Format(_T("%f"),lpProp->Value.MVflt.lpflt[ulMVRow]);// STRING_OK
		break;
	case(PT_MV_STRING8):
		//CString overloads '=' to handle conversions
		if (lpProp->Value.MVszA.lppszA[ulMVRow] && '\0' != lpProp->Value.MVszA.lppszA[ulMVRow])
		{
			szTmp = lpProp->Value.MVszA.lppszA[ulMVRow];

			HRESULT hRes = S_OK;
			SBinary sBin = {0};
			WC_H(StringCbLengthA(lpProp->Value.MVszA.lppszA[ulMVRow],STRSAFE_MAX_CCH * sizeof(char),(size_t*)&sBin.cb));
			sBin.lpb = (LPBYTE) lpProp->Value.MVszA.lppszA[ulMVRow];
			szAltTmp = BinToHexString(&sBin,true);
		}
		break;
	case(PT_MV_UNICODE):
		//CString overloads '=' to handle conversions
		if (lpProp->Value.MVszW.lppszW[ulMVRow] && L'\0' !=lpProp->Value.MVszW.lppszW[ulMVRow])
		{
			szTmp = lpProp->Value.MVszW.lppszW[ulMVRow];

			HRESULT hRes = S_OK;
			SBinary sBin = {0};
			WC_H(StringCbLengthW(lpProp->Value.MVszW.lppszW[ulMVRow],STRSAFE_MAX_CCH * sizeof(WCHAR),(size_t*)&sBin.cb));
			sBin.lpb = (LPBYTE) lpProp->Value.MVszW.lppszW[ulMVRow];
			szAltTmp = BinToHexString(&sBin,true);
		}
		break;
	case(PT_MV_BINARY):
		szTmp = BinToHexString(&lpProp->Value.MVbin.lpbin[ulMVRow],true);
		szAltTmp = BinToTextString(&lpProp->Value.MVbin.lpbin[ulMVRow],false);
		break;
	case(PT_MV_CLSID):
		{
			LPTSTR szGuid = GUIDToStringAndName(&lpProp->Value.MVguid.lpguid[ulMVRow]);
			szTmp = szGuid;
			delete[] szGuid;
		}
		break;
	default:
		break;
	}
	if (PropString) *PropString = szTmp;
	if (AltPropString) *AltPropString = szAltTmp;
}

/***************************************************************************
 Name	  : InterpretProp
 Purpose	: Evaluate a property value and return a string representing the property.
 Parameters:
 IN:
	LPSPropValue lpProp: Property to be evaluated
 OUT:
	CString *tmpPropString: String representing property value
	CString *tmpAltPropString: Alternative string representation
 Comment	: Add new Property ID's as they become known
***************************************************************************/
void InterpretProp(LPSPropValue lpProp, CString *PropString, CString *AltPropString)
{
	CString tmpPropString;
	CString tmpAltPropString;
	CString szTmp;
	CString	szAltTmp;
	ULONG iMVCount = 0;

	if (!lpProp) return;

	if (MV_FLAG & PROP_TYPE(lpProp->ulPropTag))
	{
		//MV property
		//All the MV structures are basically the same, so we can cheat when we pull the count
		ULONG cValues = lpProp->Value.MVi.cValues;
		tmpPropString.Format(_T("%d: "),cValues);// STRING_OK
		for (iMVCount = 0; iMVCount < cValues; iMVCount++)
		{
			if (iMVCount != 0)
			{
				tmpPropString += _T("; ");// STRING_OK
				switch(PROP_TYPE(lpProp->ulPropTag))
				{
				case(PT_MV_LONG):
				case(PT_MV_BINARY):
				case(PT_MV_SYSTIME):
				case(PT_MV_STRING8):
				case(PT_MV_UNICODE):
					tmpAltPropString += _T("; ");// STRING_OK
					break;
				}
			}
			InterpretMVProp(lpProp, iMVCount, &szTmp, &szAltTmp);
			tmpPropString += szTmp;
			tmpAltPropString += szAltTmp;
		}
	}
	else
	{
		switch(PROP_TYPE(lpProp->ulPropTag))
		{
		case(PT_I2):
			tmpPropString.Format(_T("%d"),lpProp->Value.i);// STRING_OK
			tmpAltPropString.Format(_T("0x%X"),lpProp->Value.i);// STRING_OK
			break;
		case(PT_LONG):
			tmpPropString.Format(_T("%u"),lpProp->Value.l);// STRING_OK
			tmpAltPropString.Format(_T("0x%X"),lpProp->Value.l);// STRING_OK
			break;
		case(PT_R4):
			tmpPropString.Format(_T("%f"),lpProp->Value.flt);// STRING_OK
			break;
		case(PT_DOUBLE):
			tmpPropString.Format(_T("%f"),lpProp->Value.dbl);// STRING_OK
			break;
		case(PT_CURRENCY):
			tmpPropString = CurrencyToString(lpProp->Value.cur);
			tmpAltPropString.Format(_T("0x%08X:0x%08X"),(int)(lpProp->Value.cur.Hi),(int)lpProp->Value.cur.Lo);// STRING_OK
			break;
		case(PT_APPTIME):
			tmpPropString.Format(_T("%u"),lpProp->Value.at);// STRING_OK
			break;
		case(PT_ERROR):
			tmpPropString.Format(_T("Err:0x%08X=%s"),lpProp->Value.err,ErrorNameFromErrorCode(lpProp->Value.err));// STRING_OK
			break;
		case(PT_BOOLEAN):
			if (lpProp->Value.b)
				tmpPropString.FormatMessage(IDS_TRUE);
			else
				tmpPropString.FormatMessage(IDS_FALSE);
			break;
		case(PT_OBJECT):
			tmpPropString.FormatMessage(IDS_OBJECT);
			break;
		case(PT_I8)://LARGE_INTEGER
			tmpPropString.Format(_T("%I64d"),lpProp->Value.li.QuadPart);// STRING_OK
			tmpAltPropString.Format(_T("0x%08X:0x%08X"),(int)(lpProp->Value.li.HighPart),(int)lpProp->Value.li.LowPart);// STRING_OK
			break;
		case(PT_STRING8):
			//CString overloads '=' to handle conversions
			if (CheckStringProp(lpProp,PT_STRING8))
			{
				tmpPropString = lpProp->Value.lpszA;

				HRESULT hRes = S_OK;
				SBinary sBin = {0};
				WC_H(StringCbLengthA(lpProp->Value.lpszA,STRSAFE_MAX_CCH * sizeof(char),(size_t*)&sBin.cb));
				sBin.lpb = (LPBYTE) lpProp->Value.lpszA;
				tmpAltPropString = BinToHexString(&sBin,true);
			}
			break;
		case(PT_UNICODE):
			//CString overloads '=' to handle conversions
			if (CheckStringProp(lpProp,PT_UNICODE))
			{
				tmpPropString = lpProp->Value.lpszW;

				HRESULT hRes = S_OK;
				SBinary sBin = {0};
				WC_H(StringCbLengthW(lpProp->Value.lpszW,STRSAFE_MAX_CCH * sizeof(WCHAR),(size_t*)&sBin.cb));
				sBin.lpb = (LPBYTE) lpProp->Value.lpszW;
				tmpAltPropString = BinToHexString(&sBin,true);
			}
			break;
		case(PT_SYSTIME):
			FileTimeToString(&lpProp->Value.ft,&tmpPropString,&tmpAltPropString);
			break;
		case(PT_CLSID):
			{
				LPTSTR szGuid = GUIDToStringAndName(lpProp->Value.lpguid);
				tmpPropString = szGuid;
				delete[] szGuid;
			}
			break;
		case(PT_BINARY):
			tmpPropString = BinToHexString(&lpProp->Value.bin,true);
			tmpAltPropString = BinToTextString(&lpProp->Value.bin,true);
			break;
		case(PT_SRESTRICTION):
			tmpPropString = RestrictionToString((LPSRestriction) lpProp->Value.lpszA,NULL);
			break;
		case(PT_ACTIONS):
			ActionsToString((ACTIONS*)lpProp->Value.lpszA,&tmpPropString);
			break;
		default:
			break;
		}
	}
	if (PropString) *PropString = tmpPropString;
	if (AltPropString) *AltPropString = tmpAltPropString;
}//InterpretProp

CString TypeToString(ULONG ulPropTag)
{
	CString tmpPropType;

	bool bNeedInstance = false;
	if (ulPropTag & MV_INSTANCE)
	{
		ulPropTag &= ~MV_INSTANCE;
		bNeedInstance = true;
	}

	ULONG ulCur = 0;
	BOOL bTypeFound = false;

	for (ulCur = 0 ; ulCur < ulPropTypeArray ; ulCur++)
	{
		if (PropTypeArray[ulCur].ulValue == PROP_TYPE(ulPropTag))
		{
			tmpPropType = PropTypeArray[ulCur].lpszName;
			bTypeFound = true;
			break;
		}
	}
	if (!bTypeFound)
		tmpPropType.Format(_T("0x%04x"),PROP_TYPE(ulPropTag));// STRING_OK

	if (bNeedInstance) tmpPropType += _T(" | MV_INSTANCE");// STRING_OK
	return tmpPropType;
}//TypeToString

//Allocates lpszPropName and lpszPropGUID with new
//Free with delete[]
void GetPropName(LPMAPIPROP lpMAPIProp,
				 ULONG ulPropTag,
				 LPTSTR *lpszPropName,
				 LPTSTR *lpszPropGUID)
{
	GetPropName(lpMAPIProp,
		ulPropTag,
		lpszPropName,
		lpszPropGUID,
		NULL);
}


// Allocates lpszPropName lpszPropGUID and lpszDASL with new
// Free with delete[]
//
// lpszDASL string for a named prop will look like this:
// id/{12345678-1234-1234-1234-12345678ABCD}/80010003
// string/{12345678-1234-1234-1234-12345678ABCD}/MyProp
// So the following #defines give the size of the buffers we need, in TCHARS, including 1 for the null terminator
// CCH_DASL_ID gets an extra digit to handle some AB props with name IDs of five digits
#define CCH_DASL_ID 2+1+38+1+8+1+1
#define CCH_DASL_STRING 6+1+38+1+1
// TagToString will prepend the http://schemas.microsoft.com/MAPI/ for us since it's a constant
// We don't compute a DASL string for non-named props as FormatMessage in TagToString can handle those

void GetPropName(LPMAPIPROP lpMAPIProp,
				 ULONG ulPropTag,
				 LPTSTR *lpszPropName,
				 LPTSTR *lpszPropGUID,
				 LPTSTR *lpszDASL)
{
	if (lpszPropName) *lpszPropName = NULL;
	if (lpszPropGUID) *lpszPropGUID = NULL;
	if (lpszDASL) *lpszDASL = NULL;

	if (!lpMAPIProp) return;
	if (!RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD) return;
	if (!RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD && PROP_ID(ulPropTag) < 0x8000) return;

	HRESULT			hRes = S_OK;
	SPropTagArray	tag = {0};
	LPSPropTagArray	lpTag = &tag;
	ULONG			ulPropNames = 0;
	LPMAPINAMEID*	lppPropNames = 0;

	tag.cValues = 1;
	tag.aulPropTag[0] = ulPropTag;

	WC_H_GETPROPS(lpMAPIProp->GetNamesFromIDs(
		&lpTag,
		NULL,
		NULL,
		&ulPropNames,
		&lppPropNames));
	if (SUCCEEDED(hRes) && ulPropNames == 1 && lppPropNames && lppPropNames[0])
	{
		DebugPrint(DBGNamedProp,_T("Got a named prop array\n"));
		DebugPrint(DBGNamedProp,_T("lpTag->aulPropTag[0] = 0x%08x\n"),lpTag->aulPropTag[0]);
		LPTSTR szGuid = GUIDToStringAndName(lppPropNames[0]->lpguid);
		DebugPrint(DBGNamedProp,_T("lppPropNames[0]->lpguid = %s\n"), szGuid);

		if (lpszPropGUID)
		{
			*lpszPropGUID = szGuid;
		}
		else
		{
			delete[] szGuid;//if we're not giving the string back, then we need to clean it up
			szGuid = NULL;
		}

		LPTSTR szDASLGuid = NULL;
		if (lpszDASL) szDASLGuid = GUIDToString(lppPropNames[0]->lpguid);

		if (lppPropNames[0]->ulKind == MNID_ID)
		{
			DebugPrint(DBGNamedProp,_T("lppPropNames[0]->Kind.lID = 0x%04X = %u\n"),lppPropNames[0]->Kind.lID,lppPropNames[0]->Kind.lID);
			if (lpszPropName)
			{
				LPCWSTR szName = NameIDToPropName(lppPropNames[0]);

				if (szName)
				{
					size_t cchName = 0;
					EC_H(StringCchLengthW(szName,STRSAFE_MAX_CCH,&cchName));
					if (SUCCEEDED(hRes))
					{
						// Worst case is 'id: 0xFFFFFFFF=4294967295' - 26 chars
						*lpszPropName = new TCHAR[26 + 3 + cchName + 1];
						if (*lpszPropName)
						{
							// Printing hex first gets a nice sort without spacing tricks
							EC_H(StringCchPrintf(*lpszPropName,26 + 3 + cchName + 1,_T("id: 0x%04X=%u = %ws"),// STRING_OK
								lppPropNames[0]->Kind.lID,
								lppPropNames[0]->Kind.lID,
								szName));
						}
					}
				}
				else
				{
					// Worst case is 'id: 0xFFFFFFFF=4294967295' - 26 chars
					*lpszPropName = new TCHAR[26];
					if (*lpszPropName)
					{
						// Printing hex first gets a nice sort without spacing tricks
						EC_H(StringCchPrintf(*lpszPropName,26 + 1,_T("id: 0x%04X=%u"),// STRING_OK
							lppPropNames[0]->Kind.lID,
							lppPropNames[0]->Kind.lID));
					}
				}
			}
			if (lpszDASL)
			{
				*lpszDASL = new TCHAR[CCH_DASL_ID];
				if (*lpszDASL)
				{
					EC_H(StringCchPrintf(*lpszDASL,CCH_DASL_ID,_T("id/%s/%04X%04X"),// STRING_OK
						szDASLGuid,
						lppPropNames[0]->Kind.lID,
						PROP_TYPE(ulPropTag)));
				}
			}
		}
		else if (lppPropNames[0]->ulKind == MNID_STRING)
		{
			//lpwstrName is LPWSTR which means it's ALWAYS unicode
			//But some folks get it wrong and stuff ANSI data in there
			//So we check the string length both ways to make our best guess
			size_t cchShortLen = strlen((LPSTR)(lppPropNames[0]->Kind.lpwstrName));
			size_t cchWideLen = wcslen(lppPropNames[0]->Kind.lpwstrName);

			if (cchShortLen < cchWideLen)
			{
				//this is the *proper* case
				DebugPrint(DBGNamedProp,_T("lppPropNames[0]->Kind.lpwstrName = \"%ws\"\n"),lppPropNames[0]->Kind.lpwstrName);
				if (lpszPropName)
				{
					*lpszPropName = new TCHAR[7+cchWideLen];
					if (*lpszPropName)
					{
//Compiler Error C2017 - Can occur (falsly) when escape sequences are stringized, as EC_H will do here
#define __GOODSTRING _T("sz: \"%ws\"") // STRING_OK
						EC_H(StringCchPrintf(*lpszPropName,7+cchWideLen,__GOODSTRING,
							lppPropNames[0]->Kind.lpwstrName));
					}
				}
				if (lpszDASL)
				{
					*lpszDASL = new TCHAR[CCH_DASL_STRING+cchWideLen];
					if (*lpszDASL)
					{
						EC_H(StringCchPrintf(*lpszDASL,CCH_DASL_STRING +cchWideLen,_T("string/%s/%ws"),// STRING_OK
							szDASLGuid,
							lppPropNames[0]->Kind.lpwstrName));
					}
				}
			}
			else
			{
				//this is the case where ANSI data was shoved into a unicode string.
				DebugPrint(DBGNamedProp,_T("Warning: ANSI data was found in a unicode field. This is a bug on the part of the creator of this named property\n"));
				DebugPrint(DBGNamedProp,_T("lppPropNames[0]->Kind.lpwstrName = \"%hs\"\n"),lppPropNames[0]->Kind.lpwstrName);
				if (lpszPropName)
				{
					*lpszPropName = new TCHAR[7+cchShortLen+25];
					if (*lpszPropName)
					{
						CString szComment;
						szComment.LoadString(IDS_NAMEWASANSI);
//Compiler Error C2017 - Can occur (falsly) when escape sequences are stringized, as EC_H will do here
#define __BADSTRING _T("sz: \"%hs\" %s") // STRING_OK
						EC_H(StringCchPrintf(*lpszPropName,7+cchShortLen+25,__BADSTRING,
							lppPropNames[0]->Kind.lpwstrName,szComment));
					}
				}
				if (lpszDASL)
				{
					*lpszDASL = new TCHAR[CCH_DASL_STRING+cchShortLen];
					if (*lpszDASL)
					{
						EC_H(StringCchPrintf(*lpszDASL,CCH_DASL_STRING+cchShortLen,_T("string/%s/%hs"),// STRING_OK
							szDASLGuid,
							lppPropNames[0]->Kind.lpwstrName));
					}
				}
			}
		}
		delete[] szDASLGuid;
	}
	MAPIFreeBuffer(lppPropNames);
}