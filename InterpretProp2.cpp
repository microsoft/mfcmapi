#include "stdafx.h"
#include "Error.h"

#include "InterpretProp2.h"
#include "InterpretProp.h"
#include "PropTagArray.h"
#include "ExtraPropTags.h"
#include "MAPIFunctions.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define ulNoMatch 0xffffffff
static WCHAR szPropSeparator[] = L", ";// STRING_OK

// lpszExactMatch and lpszPartialMatches allocated with new
// clean up with delete[]
HRESULT PropTagToPropName(ULONG ulPropTag, BOOL bIsAB, LPTSTR* lpszExactMatch, LPTSTR* lpszPartialMatches)
{
	if (!lpszExactMatch && !lpszPartialMatches) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	ULONG ulPropID = NULL;
	ULONG ulPropType = NULL;
	ULONG ulLowerBound = 0;
	ULONG ulUpperBound = ulPropTagArray-1;// ulPropTagArray-1 is the last entry
	ULONG ulMidPoint = (ulUpperBound+ulLowerBound)/2;
	ULONG ulFirstMatch = ulNoMatch;
	ULONG ulLastMatch = ulNoMatch;
	ULONG ulFirstExactMatch = ulNoMatch;
	ULONG ulLastExactMatch = ulNoMatch;

	if (lpszExactMatch) *lpszExactMatch = NULL;
	if (lpszPartialMatches) *lpszPartialMatches = NULL;

	if (!ulPropTagArray || !PropTagArray) return S_OK;

	//determine the prop ID we're seeking
	if (ulPropTag & 0xffff0000)//dealing with a full prop tag
	{
		ulPropID   = PROP_ID(ulPropTag);
		ulPropType = PROP_TYPE(ulPropTag);
	}
	else
	{
		ulPropID   = ulPropTag;
		ulPropType = PT_UNSPECIFIED;
	}

	//short circuit property IDs with the high bit set if bIsAB wasn't passed
	if (!bIsAB && (ulPropID & 0x8000)) return hRes;

	//put everything back together
	ulPropTag = PROP_TAG(ulPropType,ulPropID);

	//find A match
	while (ulUpperBound - ulLowerBound > 1)
	{
		if (ulPropID == PROP_ID(PropTagArray[ulMidPoint].ulValue))
		{
			ulFirstMatch = ulMidPoint;
			break;
		}

		if (ulPropID < PROP_ID(PropTagArray[ulMidPoint].ulValue))
		{
			ulUpperBound = ulMidPoint;
		}
		else if (ulPropID > PROP_ID(PropTagArray[ulMidPoint].ulValue))
		{
			ulLowerBound = ulMidPoint;
		}
		ulMidPoint = (ulUpperBound+ulLowerBound)/2;
	}

	//when we get down to two points, we may have only checked one of them
	//make sure we've checked the other
	if (ulPropID == PROP_ID(PropTagArray[ulUpperBound].ulValue))
	{
		ulFirstMatch = ulUpperBound;
	}
	else if (ulPropID == PROP_ID(PropTagArray[ulLowerBound].ulValue))
	{
		ulFirstMatch = ulLowerBound;
	}

	// check that we got a match
	if (ulNoMatch != ulFirstMatch)
	{
		ulLastMatch = ulFirstMatch;//remember the last match we've found so far

		//scan backwards to find the first match
		while (ulFirstMatch > 0 && ulPropID == PROP_ID(PropTagArray[ulFirstMatch-1].ulValue))
		{
			ulFirstMatch = ulFirstMatch - 1;
		}

		//scan forwards to find the real last match
		// Last entry in the array is ulPropTagArray-1
		while (ulLastMatch+1 < ulPropTagArray && ulPropID == PROP_ID(PropTagArray[ulLastMatch+1].ulValue))
		{
			ulLastMatch = ulLastMatch + 1;
		}

		//scan to see if we have any exact matches
		ULONG ulCur;
		for (ulCur = ulFirstMatch ; ulCur <= ulLastMatch ; ulCur++)
		{
			if (ulPropTag == PropTagArray[ulCur].ulValue)
			{
				ulFirstExactMatch = ulCur;
				break;
			}
		}

		ULONG ulNumExacts = 0;

		if (ulNoMatch != ulFirstExactMatch)
		{
			for (ulCur = ulFirstExactMatch ; ulCur <= ulLastMatch ; ulCur++)
			{
				if (ulPropTag == PropTagArray[ulCur].ulValue)
				{
					ulLastExactMatch = ulCur;
				}
				else break;
			}
			ulNumExacts = ulLastExactMatch - ulFirstExactMatch + 1;
		}

		if (lpszExactMatch && ulNumExacts > 0 && ulNoMatch != ulFirstExactMatch)
		{
			size_t cchExact = 1 + (ulNumExacts - 1) * (sizeof(szPropSeparator)/sizeof(TCHAR)-1);
			for (ulCur = ulFirstExactMatch ; ulCur <= ulLastExactMatch ; ulCur++)
			{
				size_t cchLen = 0;
				EC_H(StringCchLengthW(PropTagArray[ulCur].lpszName,STRSAFE_MAX_CCH,&cchLen));
				cchExact += cchLen;
			}

			LPWSTR szExactMatch = NULL;
			szExactMatch = new WCHAR[cchExact];
			if (szExactMatch)
			{
				szExactMatch[0] = _T('\0');
				for (ulCur = ulFirstExactMatch ; ulCur <= ulLastExactMatch ; ulCur++)
				{
					EC_H(StringCchCatW(szExactMatch,cchExact,PropTagArray[ulCur].lpszName));
					if (ulCur < ulLastExactMatch)
					{
						EC_H(StringCchCatW(szExactMatch,cchExact,szPropSeparator));
					}
				}
				if (SUCCEEDED(hRes))
				{
#ifdef UNICODE
					*lpszExactMatch = szExactMatch;
#else
					LPSTR szAnsiExactMatch = NULL;
					EC_H(UnicodeToAnsi(szExactMatch,&szAnsiExactMatch));
					if (SUCCEEDED(hRes))
					{
						*lpszExactMatch = szAnsiExactMatch;
					}
					delete[] szExactMatch;
#endif
				}
			}
		}

		ULONG ulNumPartials = ulLastMatch - ulFirstMatch + 1 - ulNumExacts;

		if (lpszPartialMatches && ulNumPartials > 0 && lpszPartialMatches)
		{
			//let's build lpszPartialMatches
			//see how much space we need - initialize cchPartial with space for separators and NULL terminator
			//note - ulNumPartials-1 is the number of spaces we need...
			size_t cchPartial = 1 + (ulNumPartials - 1) * (sizeof(szPropSeparator)/sizeof(TCHAR)-1);
			for (ulCur = ulFirstMatch ; ulCur <= ulLastMatch ; ulCur++)
			{
				if (ulPropTag == PropTagArray[ulCur].ulValue) continue;//skip our exact matches
				size_t cchLen = 0;
				EC_H(StringCchLengthW(PropTagArray[ulCur].lpszName,STRSAFE_MAX_CCH,&cchLen));
				cchPartial += cchLen;
			}

			LPWSTR szPartialMatches = NULL;
			szPartialMatches = new WCHAR[cchPartial];
			if (szPartialMatches)
			{
				szPartialMatches[0] = _T('\0');
				ULONG ulNumSeparators = 1;//start at 1 so we print one less than we print strings
				for (ulCur = ulFirstMatch ; ulCur <= ulLastMatch ; ulCur++)
				{
					if (ulPropTag == PropTagArray[ulCur].ulValue) continue;//skip our exact matches
					EC_H(StringCchCatW(szPartialMatches,cchPartial,PropTagArray[ulCur].lpszName));
					if (ulNumSeparators < ulNumPartials)
					{
						EC_H(StringCchCatW(szPartialMatches,cchPartial,szPropSeparator));
					}
					ulNumSeparators++;
				}
				if (SUCCEEDED(hRes))
				{
#ifdef UNICODE
					*lpszPartialMatches = szPartialMatches;
#else
					LPSTR szAnsiPartialMatches = NULL;
					EC_H(UnicodeToAnsi(szPartialMatches,&szAnsiPartialMatches));
					if (SUCCEEDED(hRes))
					{
						*lpszPartialMatches = szAnsiPartialMatches;
					}
					delete[] szPartialMatches;
#endif
				}
			}
		}
	}

	return hRes;
}

HRESULT PropNameToPropTag(LPCTSTR lpszPropName, ULONG* ulPropTag)
{
	if (!lpszPropName || !ulPropTag) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	ULONG ulCur = 0;

	*ulPropTag = NULL;
	if (!ulPropTagArray || !PropTagArray) return S_OK;

#ifdef UNICODE
	LPCWSTR szPropName = lpszPropName;
#else
	LPWSTR szPropName = NULL;
	EC_H(AnsiToUnicode(lpszPropName,&szPropName));
	if (SUCCEEDED(hRes))
	{
#endif

	for (ulCur = 0 ; ulCur < ulPropTagArray ; ulCur++)
	{
		if (0 == lstrcmpiW(szPropName,PropTagArray[ulCur].lpszName))
		{
			*ulPropTag = PropTagArray[ulCur].ulValue;
			break;
		}
	}

#ifndef UNICODE
	}
	delete[] szPropName;
#endif
	return hRes;
}

HRESULT PropTypeNameToPropType(LPCTSTR lpszPropType, ULONG* ulPropType)
{
	if (!lpszPropType || !ulPropType) return MAPI_E_INVALID_PARAMETER;

	ULONG ulCur = 0;

	*ulPropType = NULL;
	if (!ulPropTypeArray || !PropTypeArray) return S_OK;

#ifdef UNICODE
	LPCWSTR szPropType = lpszPropType;
#else
	HRESULT hRes = S_OK;
	LPWSTR szPropType = NULL;
	EC_H(AnsiToUnicode(lpszPropType,&szPropType));
	if (SUCCEEDED(hRes))
	{
#endif

	for (ulCur = 0 ; ulCur < ulPropTypeArray ; ulCur++)
	{
		if (0 == lstrcmpiW(szPropType,PropTypeArray[ulCur].lpszName))
		{
			*ulPropType = PropTypeArray[ulCur].ulValue;
			break;
		}
	}

#ifndef UNICODE
	}
	delete[] szPropType;
#endif
	return S_OK;
}

LPTSTR GUIDToStringAndName(LPCGUID lpGUID)
{
	HRESULT	hRes = S_OK;
	ULONG	ulCur = 0;
	LPCWSTR	szGUIDName = NULL;
	size_t	cchGUIDName = NULL;
	size_t	cchGUID = NULL;
	WCHAR	szUnknown[13]; // The length of IDS_UNKNOWNGUID

	if (lpGUID && ulPropGuidArray && PropGuidArray)
	{
		for (ulCur = 0 ; ulCur < ulPropGuidArray ; ulCur++)
		{
			if (IsEqualGUID(*lpGUID,*PropGuidArray[ulCur].lpGuid))
			{
				szGUIDName = PropGuidArray[ulCur].lpszName;
				break;
			}
		}
	}
	if (!szGUIDName)
	{
		int iRet = NULL;
		// CString doesn't provide a way to extract just Unicode strings, so we do this manually
		EC_D(iRet,LoadStringW(GetModuleHandle(NULL),
			IDS_UNKNOWNGUID,
			szUnknown,
			CCHW(szUnknown)));
		szGUIDName = szUnknown;
	}

	LPTSTR szGUID = GUIDToString(lpGUID);

	EC_H(StringCchLengthW(szGUIDName,STRSAFE_MAX_CCH,&cchGUIDName));
	if (szGUID) EC_H(StringCchLength(szGUID,STRSAFE_MAX_CCH,&cchGUID));

	size_t cchBothGuid = cchGUID + 3 + cchGUIDName+1;
	LPTSTR szBothGuid = new TCHAR[cchBothGuid];

	if (szBothGuid)
	{
		EC_H(StringCchPrintf(szBothGuid,cchBothGuid,_T("%s = %ws"),szGUID,szGUIDName));// STRING_OK
	}

	delete[] szGUID;
	return szBothGuid;
}

void GUIDNameToGUID(LPCTSTR szGUID, LPCGUID* lpGUID)
{
	if (!szGUID || !lpGUID) return;

	ULONG ulCur = 0;

	*lpGUID = NULL;

	if (!ulPropGuidArray || !PropGuidArray) return;

#ifdef UNICODE
	LPCWSTR szGUIDW = szGUID;
#else
	HRESULT hRes = S_OK;
	LPWSTR szGUIDW = NULL;
	EC_H(AnsiToUnicode(szGUID,&szGUIDW));
	if (SUCCEEDED(hRes))
	{
#endif
	for (ulCur = 0 ; ulCur < ulPropGuidArray ; ulCur++)
	{
		if (0 == lstrcmpiW(szGUIDW,PropGuidArray[ulCur].lpszName))
		{
			*lpGUID = PropGuidArray[ulCur].lpGuid;
			break;
		}
	}
#ifndef UNICODE
	}
	delete[] szGUIDW;
#endif
}

// Returns string from NameIDArray - do not delete
LPCWSTR NameIDToPropName(LPMAPINAMEID lpNameID)
{
	if (!lpNameID) return NULL;
	if (lpNameID->ulKind != MNID_ID) return NULL;
	ULONG	ulCur = 0;
	ULONG	ulMatch = ulNoMatch;

	if (!ulNameIDArray || !NameIDArray) return NULL;

	for (ulCur = 0 ; ulCur < ulNameIDArray ; ulCur++)
	{
		if (NameIDArray[ulCur].lValue == lpNameID->Kind.lID)
		{
			ulMatch = ulCur;
			break;
		}
	}
	if (ulNoMatch == ulMatch)
	{
		return NULL;
	}

	for (ulCur = ulMatch ; ulCur < ulNameIDArray ; ulCur++)
	{
		if (NameIDArray[ulCur].lValue != lpNameID->Kind.lID)
		{
			return NULL;
		}
		if (IsEqualGUID(*lpNameID->lpguid,*NameIDArray[ulCur].lpGuid))
		{
			return NameIDArray[ulCur].lpszName;
		}
	}
	return NULL;
}// NameIDToPropName

//Interprets a flag found in lpProp and returns a string allocated with MAPIAllocateBuffer
//Free the string with MAPIFreeBuffer
//Will not return a string if the lpProp is not a PT_LONG or we don't recognize the property
HRESULT InterpretFlags(LPSPropValue lpProp, LPTSTR* szFlagString)
{
	if (szFlagString) *szFlagString = NULL;
	if (!lpProp || PROP_TYPE(lpProp->ulPropTag) != PT_LONG || !szFlagString)
	{
		return S_OK;//remove this if we encounter other types needing interpreting
	}

	return InterpretFlags(PROP_ID(lpProp->ulPropTag),lpProp->Value.ul,szFlagString);
}

//Interprets a flag value according to a flag name and returns a string
//   allocated with MAPIAllocateBuffer
//Free the string with MAPIFreeBuffer
//Will not return a string if the flag name is not recognized
HRESULT InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, LPTSTR* szFlagString)
{
	if (!szFlagString)
	{
		return S_OK;
	}
	HRESULT	hRes = S_OK;
	ULONG	ulCurEntry = 0;
	LONG	lTempValue = lFlagValue;
	WCHAR	szTempString[1024];

	*szFlagString = NULL;
	szTempString[0] = NULL;

	if (!ulFlagArray || !FlagArray) return S_OK;

	while (ulCurEntry < ulFlagArray && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
	{
		ulCurEntry++;
	}

	if (FlagArray[ulCurEntry].ulFlagName != ulFlagName) return S_OK;

	//We've matched our flag name to the array - we SHOULD return a string at this point
	BOOL bNeedSeparator = false;

	for (;FlagArray[ulCurEntry].ulFlagName == ulFlagName;ulCurEntry++)
	{
		if (flagFLAG == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue & lTempValue)
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,CCHW(szTempString),L" | "));// STRING_OK
				}
				EC_H(StringCchCatW(szTempString,CCHW(szTempString),FlagArray[ulCurEntry].lpszName));
				lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
				bNeedSeparator = true;
			}
		}
		else if (flagVALUE == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue == lTempValue)
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,CCHW(szTempString),L" | "));// STRING_OK
				}
				EC_H(StringCchCatW(szTempString,CCHW(szTempString),FlagArray[ulCurEntry].lpszName));
				lTempValue = 0;
				bNeedSeparator = true;
			}
		}
		else if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
		{
			//find any bits we need to clear
			LONG lClearedBits = FlagArray[ulCurEntry].lFlagValue & lTempValue;
			//report what we found
			if (0 != lClearedBits)
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,CCHW(szTempString),L" | "));// STRING_OK
				}
				WCHAR szClearedBits[15];
				EC_H(StringCchPrintfW(szClearedBits,CCHW(szClearedBits),L"0x%08X",lClearedBits));// STRING_OK
				EC_H(StringCchCatW(szTempString,CCHW(szTempString),szClearedBits));
				//clear the bits out
				lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
				bNeedSeparator = true;
			}
		}
	}

	// We know if we've found anything already because bNeedSeparator will be true
	// If bNeedSeparator isn't true, we found nothing and need to tack on
	// Otherwise, it's true, and we only tack if lTempValue still has something in it
	if (!bNeedSeparator || lTempValue)
	{
		WCHAR	szUnk[15];
		if (bNeedSeparator)
		{
			EC_H(StringCchCatW(szTempString,CCHW(szTempString),L" | "));// STRING_OK
		}
		EC_H(StringCchPrintfW(szUnk,CCHW(szUnk),L"0x%08X",lTempValue));// STRING_OK
		EC_H(StringCchCatW(szTempString,CCHW(szTempString),szUnk));
	}

	//Copy the string we computed for output
	size_t cchLen = 0;
	EC_H(StringCchLengthW(szTempString,CCHW(szTempString),&cchLen));

	if (cchLen)
	{
		cchLen++;//for the NULL
		EC_H(MAPIAllocateBuffer(
			(ULONG) (sizeof(TCHAR) * cchLen),
			(LPVOID*) szFlagString));

		if (*szFlagString)
		{
			(*szFlagString)[0] = NULL;
			EC_H(StringCchPrintf(*szFlagString,cchLen,_T("%ws"),szTempString));// STRING_OK
		}
	}

	return S_OK;
}

// Returns a list of all known flags/values for a flag name.
// For instance, for flagFuzzyLevel, would return:
//\r\n0x00000000 FL_FULLSTRING\r\n\
//0x00000001 FL_SUBSTRING\r\n\
//0x00000002 FL_PREFIX\r\n\
//0x00010000 FL_IGNORECASE\r\n\
//0x00020000 FL_IGNORENONSPACE\r\n\
//0x00040000 FL_LOOSE
//
// Since the string is always appended to a prompt we include \r\n at the start
CString AllFlagsToString(const ULONG ulFlagName,BOOL bHex)
{
	CString szFlagString;
	if (!ulFlagName) return szFlagString;
	if (!ulFlagArray || !FlagArray) return szFlagString;

	ULONG	ulCurEntry = 0;
	CString szTempString;

	while (ulCurEntry < ulFlagArray && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
	{
		ulCurEntry++;
	}

	if (FlagArray[ulCurEntry].ulFlagName != ulFlagName) return szFlagString;

	//We've matched our flag name to the array - we SHOULD return a string at this point
	for (;FlagArray[ulCurEntry].ulFlagName == ulFlagName;ulCurEntry++)
	{
		if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
		{
			// keep going
		}
		else
		{
			if (bHex)
			{
				szTempString.FormatMessage(_T("\r\n0x%1!08X! %2!ws!"),FlagArray[ulCurEntry].lFlagValue,FlagArray[ulCurEntry].lpszName);// STRING_OK
			}
			else
			{
				szTempString.FormatMessage(_T("\r\n%1!5d! %2!ws!"),FlagArray[ulCurEntry].lFlagValue,FlagArray[ulCurEntry].lpszName);// STRING_OK
			}
			szFlagString += szTempString;
		}
	}

	return szFlagString;
}
