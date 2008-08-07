#include "stdafx.h"
#include "Error.h"

#include "InterpretProp2.h"
#include "InterpretProp.h"
#include "PropTagArray.h"
#include "ExtraPropTags.h"
#include "MAPIFunctions.h"
#include "guids.h"
#include "MySecInfo.h"

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
			size_t cchExact = 1 + (ulNumExacts - 1) * (sizeof(szPropSeparator)/sizeof(WCHAR)-1);
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
			size_t cchPartial = 1 + (ulNumPartials - 1) * (sizeof(szPropSeparator)/sizeof(WCHAR)-1);
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

// Allocates and returns string built from NameIDArray
// Allocated with new, clean up with delete[]
LPCWSTR NameIDToPropName(LPMAPINAMEID lpNameID)
{
	if (!lpNameID) return NULL;
	if (lpNameID->ulKind != MNID_ID) return NULL;
	HRESULT hRes = S_OK;
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

	// count up how long our string needs to be
	size_t cchResultString = 0;
	ULONG ulNumMatches = 0;
	for (ulCur = ulMatch ; ulCur < ulNameIDArray ; ulCur++)
	{
		size_t cchLen = 0;
		if (NameIDArray[ulCur].lValue != lpNameID->Kind.lID)
		{
			break;
		}
		if (IsEqualGUID(*lpNameID->lpguid,*NameIDArray[ulCur].lpGuid))
		{
			EC_H(StringCchLengthW(NameIDArray[ulCur].lpszName,STRSAFE_MAX_CCH,&cchLen));
			cchResultString += cchLen;
			ulNumMatches++;
		}
	}

	if (!ulNumMatches) return NULL;

	// Add in space for null terminator and separators
	cchResultString += 1 + (ulNumMatches - 1) * (sizeof(szPropSeparator)/sizeof(WCHAR)-1);

	// Copy our matches into the result string
	LPWSTR szResultString = NULL;
	szResultString = new WCHAR[cchResultString];
	if (szResultString)
	{
		szResultString[0] = L'\0'; // STRING_OK
		for (ulCur = ulMatch ; ulCur < ulNameIDArray ; ulCur++)
		{
			if (NameIDArray[ulCur].lValue != lpNameID->Kind.lID)
			{
				break;
			}
			if (IsEqualGUID(*lpNameID->lpguid,*NameIDArray[ulCur].lpGuid))
			{
				EC_H(StringCchCatW(szResultString,cchResultString,NameIDArray[ulCur].lpszName));
				if (--ulNumMatches > 0)
				{
					EC_H(StringCchCatW(szResultString,cchResultString,szPropSeparator));
				}
			}
		}
		if (SUCCEEDED(hRes))
		{
			return szResultString;
		}
		else
		{
			delete[] szResultString;
		}
	}

	return NULL;
}// NameIDToPropName

//Interprets a flag found in lpProp and returns a string allocated with new
//Free the string with delete[]
//Will not return a string if the lpProp is not a PT_LONG/PT_I2 or we don't recognize the property
HRESULT InterpretFlags(LPSPropValue lpProp, LPTSTR* szFlagString)
{
	if (szFlagString) *szFlagString = NULL;
	if (!lpProp || !szFlagString)
	{
		return S_OK;
	}
	if (PROP_TYPE(lpProp->ulPropTag) == PT_LONG)
		return InterpretFlags(PROP_ID(lpProp->ulPropTag),lpProp->Value.ul,szFlagString);
	if (PROP_TYPE(lpProp->ulPropTag) == PT_I2)
		return InterpretFlags(PROP_ID(lpProp->ulPropTag),lpProp->Value.i,szFlagString);

	return S_OK;
}

ULONG BuildFlagIndexFromTag(ULONG ulPropTag,
							ULONG ulPropNameID,
							LPWSTR lpszPropNameString,
							LPGUID lpguidNamedProp)
{
	ULONG ulPropID = PROP_ID(ulPropTag);

	// Non-zero less than 0x8000 is a regular prop, we use the ID as the index
	if (ulPropID && ulPropID < 0x8000) return ulPropID;

	// Else we build our index from the guid and named property ID
	// In the future, we can look at using lpszPropNameString for MNID_STRING named properties
	if (lpguidNamedProp &&
		(ulPropNameID || lpszPropNameString))
	{
		ULONG ulGuid = NULL;
		if      (*lpguidNamedProp == PSETID_Meeting)        ulGuid = guidPSETID_Meeting;
		else if (*lpguidNamedProp == PSETID_Address)        ulGuid = guidPSETID_Address;
		else if (*lpguidNamedProp == PSETID_Task)           ulGuid = guidPSETID_Task;
		else if (*lpguidNamedProp == PSETID_Appointment)    ulGuid = guidPSETID_Appointment;
		else if (*lpguidNamedProp == PSETID_Common)         ulGuid = guidPSETID_Common;
		else if (*lpguidNamedProp == PSETID_Log)            ulGuid = guidPSETID_Log;
		else if (*lpguidNamedProp == PSETID_PostRss)        ulGuid = guidPSETID_PostRss;
		else if (*lpguidNamedProp == PSETID_Sharing)        ulGuid = guidPSETID_Sharing;
		else if (*lpguidNamedProp == PSETID_Note)           ulGuid = guidPSETID_Note;

		if (ulGuid && ulPropNameID)
		{
			return PROP_TAG(ulGuid,ulPropNameID);
		}
// Case not handled yet
//		else if (ulGuid && lpszPropNameString)
//		{
//		}
	}
	return NULL;
}

// Interprets a flag found in lpProp and returns a string allocated with new
// Free the string with delete[]
// Will not return a string if the lpProp is not a PT_LONG/PT_I2 or we don't recognize the property
// Will use named property details to look up named property flags
HRESULT InterpretFlags(LPSPropValue lpProp,
					   ULONG ulPropNameID,
					   LPWSTR lpszPropNameString,
					   LPGUID lpguidNamedProp,
					   LPCTSTR szPrefix,
					   LPTSTR* szFlagString)
{
	if (szFlagString) *szFlagString = NULL;
	if (!lpProp || !szFlagString)
	{
		return S_OK;
	}
	if (PROP_TYPE(lpProp->ulPropTag) != PT_LONG &&
		PROP_TYPE(lpProp->ulPropTag) != PT_I2)
	{
		return S_OK;
	}

	ULONG ulPropID = BuildFlagIndexFromTag(lpProp->ulPropTag,ulPropNameID,lpszPropNameString,lpguidNamedProp);
	LONG lVal = NULL;
	if (PROP_TYPE(lpProp->ulPropTag) == PT_LONG)
		lVal = lpProp->Value.ul;
	if (PROP_TYPE(lpProp->ulPropTag) == PT_I2)
		lVal = lpProp->Value.i;

	InterpretFlags(ulPropID,lVal,szPrefix,szFlagString);
	return S_OK;
}

// Interprets a flag value according to a flag name and returns a string
// allocated with new
// Free the string with delete[]
// Will not return a string if the flag name is not recognized
HRESULT InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, LPTSTR* szFlagString)
{
	return InterpretFlags(ulFlagName, lFlagValue, _T(""), szFlagString);
}

// Interprets a flag value according to a flag name and returns a string
// allocated with new
// Free the string with delete[]
// Will not return a string if the flag name is not recognized
HRESULT InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, LPCTSTR szPrefix, LPTSTR* szFlagString)
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

	// Don't run off the end of the array
	if (ulFlagArray == ulCurEntry) return S_OK;
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
		else if (flagVALUE3RDBYTE == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue == ((lTempValue >> 8) & 0xFF))
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,CCHW(szTempString),L" | "));// STRING_OK
				}
				EC_H(StringCchCatW(szTempString,CCHW(szTempString),FlagArray[ulCurEntry].lpszName));
				lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 8);
				bNeedSeparator = true;
			}
		}
		else if (flagVALUE4THBYTE == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0xFF))
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,CCHW(szTempString),L" | "));// STRING_OK
				}
				EC_H(StringCchCatW(szTempString,CCHW(szTempString),FlagArray[ulCurEntry].lpszName));
				lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
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
		size_t cchPrefix = NULL;
		if (szPrefix)
		{
			EC_H(StringCchLength(szPrefix,STRSAFE_MAX_CCH,&cchPrefix));
			cchLen += cchPrefix;
		}

		*szFlagString = new TCHAR[cchLen];

		if (*szFlagString)
		{
			(*szFlagString)[0] = NULL;
			EC_H(StringCchPrintf(*szFlagString,cchLen,_T("%s%ws"),szPrefix,szTempString));// STRING_OK
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

struct BINARY_STRUCTURE_ARRAY_ENTRY
{
	ULONG	ulIndex;
	MAPIStructType myStructType;
};
typedef BINARY_STRUCTURE_ARRAY_ENTRY FAR * LPBINARY_STRUCTURE_ARRAY_ENTRY;

#define BINARY_STRUCTURE_ENTRY(_fName,_fType) {PROP_ID((_fName)),(_fType)},
#define NAMEDPROP_BINARY_STRUCTURE_ENTRY(_fName,_fGuid,_fType) {PROP_TAG((guid##_fGuid),(_fName)),(_fType)},

BINARY_STRUCTURE_ARRAY_ENTRY g_BinaryStructArray[] =
{
	BINARY_STRUCTURE_ENTRY(PR_FREEBUSY_NT_SECURITY_DESCRIPTOR,stSecurityDescriptor)
	BINARY_STRUCTURE_ENTRY(PR_NT_SECURITY_DESCRIPTOR,stSecurityDescriptor)
	BINARY_STRUCTURE_ENTRY(PR_EXTENDED_FOLDER_FLAGS,stExtendedFolderFlags)
	BINARY_STRUCTURE_ENTRY(PR_REPORT_TAG,stReportTag)
	BINARY_STRUCTURE_ENTRY(PR_CONVERSATION_INDEX,stConversationIndex)

	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidTimeZoneStruct,PSETID_Appointment,stTimeZone)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidApptTZDefStartDisplay,PSETID_Appointment,stTimeZoneDefinition)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidApptTZDefEndDisplay,PSETID_Appointment,stTimeZoneDefinition)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidApptTZDefRecur,PSETID_Appointment,stTimeZoneDefinition)

	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidApptRecur,PSETID_Appointment,stAppointmentRecurrencePattern)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidTaskRecur,PSETID_Task,stRecurrencePattern)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidTaskMyDelegators,PSETID_Task,stTaskAssigners)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(LID_GLOBAL_OBJID,PSETID_Meeting,stGlobalObjectId)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(LID_CLEAN_GLOBAL_OBJID,PSETID_Meeting,stGlobalObjectId)
};

LPBINARY_STRUCTURE_ARRAY_ENTRY BinaryStructArray = g_BinaryStructArray;
ULONG ulBinaryStructArray = sizeof(g_BinaryStructArray)/sizeof(BINARY_STRUCTURE_ARRAY_ENTRY);

MAPIStructType FindStructForBinaryProp(const ULONG ulPropTag, const ULONG ulPropNameID, const LPGUID lpguidNamedProp)
{
	ULONG	ulCurEntry = 0;
	ULONG	ulIndex = BuildFlagIndexFromTag(ulPropTag,ulPropNameID,NULL,lpguidNamedProp);

	while (ulCurEntry < ulBinaryStructArray && BinaryStructArray[ulCurEntry].ulIndex != ulIndex)
	{
		ulCurEntry++;
	}

	if (BinaryStructArray[ulCurEntry].ulIndex == ulIndex) return BinaryStructArray[ulCurEntry].myStructType;
	return stUnknown;
}

// Uber property interpreter - given an LPSPropValue, produces all manner of strings
// All LPTSTR strings allocated with new, delete with delete[]
// If lpProp is NULL but ulPropTag and lpMAPIProp are passed, will call GetProps
void InterpretProp(LPSPropValue lpProp, // optional property value
				   ULONG ulPropTag, // optional 'original' prop tag
				   LPMAPIPROP lpMAPIProp, // optional source object
				   LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
				   BOOL bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
				   LPTSTR* lpszNameExactMatches, // Built from ulPropTag & bIsAB
				   LPTSTR* lpszNamePartialMatches, // Built from ulPropTag & bIsAB
				   CString* PropType, // Built from ulPropTag
				   CString* PropTag, // Built from ulPropTag
				   CString* PropString, // Built from lpProp
				   CString* AltPropString, // Built from lpProp
				   LPTSTR* lpszSmartView, // Built from lpProp & lpMAPIProp
				   LPTSTR* lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
				   LPTSTR* lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
				   LPTSTR* lpszNamedPropDASL) // Built from ulPropTag & lpMAPIProp
{
	HRESULT hRes = S_OK;

	// These four strings are based on ulPropTag, not the LPSPropValue
	if (PropType) *PropType = TypeToString(ulPropTag);
	if (lpszNameExactMatches || lpszNamePartialMatches)
		EC_H(PropTagToPropName(ulPropTag,bIsAB,lpszNameExactMatches,lpszNamePartialMatches));
	if (PropTag) PropTag->Format(_T("0x%08X"),ulPropTag);// STRING_OK

	//Named Props
	LPMAPINAMEID*	lppPropNames = 0;

	// If we weren't passed named property information and we need it, look it up
	if (!lpNameID &&
		lpMAPIProp && // if we have an object
		RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD && // and we're parsing named props
		(RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD || PROP_ID(ulPropTag) >= 0x8000) && // and it's either a named prop or we're doing all props
		(lpszNamedPropName || lpszNamedPropGUID || lpszNamedPropDASL || (lpProp && lpszSmartView))) // and we want to return something that needs named prop information
	{
		SPropTagArray	tag = {0};
		LPSPropTagArray	lpTag = &tag;
		ULONG			ulPropNames = 0;
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
			lpNameID = lppPropNames[0];
		}
		hRes = S_OK;
	}
	if (lpNameID)
	{
		NameIDToStrings(lpNameID,
			ulPropTag,
			lpszNamedPropName,
			lpszNamedPropGUID,
			lpszNamedPropDASL);
	}

	if (lpProp)
	{
		if (PropString || AltPropString)
			InterpretProp(lpProp,PropString,AltPropString);

		if (lpszSmartView)
		{
			ULONG ulPropNameID = NULL;
			if (lpNameID && lpNameID->ulKind == MNID_ID)
			{
				ulPropNameID = lpNameID->Kind.lID;
			}
			switch(PROP_TYPE(ulPropTag))
			{
			case PT_LONG:
			case PT_I2:
				{
					CString	szPrefix;
					szPrefix.LoadString(IDS_FLAGS_PREFIX);
					EC_H(InterpretFlags(lpProp,ulPropNameID,NULL,lpNameID?lpNameID->lpguid:NULL,szPrefix,lpszSmartView));
				}
				break;
			case PT_BINARY:
				{
					if (lpProp)
					{
						MAPIStructType myType = FindStructForBinaryProp(lpProp->ulPropTag,ulPropNameID,lpNameID?lpNameID->lpguid:NULL);
						if (stUnknown != myType)
						{
							InterpretBinaryAsString(lpProp->Value.bin,myType,lpMAPIProp,ulPropTag,lpszSmartView);
						}
					}
				}
				break;
			}
		}
	}
	MAPIFreeBuffer(lppPropNames);
}

// Returns LPSPropValue with value of a binary property
// Uses GetProps and falls back to OpenProperty if the value is large
// Free with MAPIFreeBuffer
HRESULT GetLargeBinaryProp(LPMAPIPROP lpMAPIProp, ULONG ulPropTag, LPSPropValue* lppProp)
{
	if (!lpMAPIProp || !lppProp) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric,_T("GetLargeBinaryProp getting buffer from 0x%08X\n"),ulPropTag);

	ulPropTag = CHANGE_PROP_TYPE(ulPropTag,PT_BINARY);

	HRESULT			hRes		= S_OK;
	ULONG			cValues		= 0;
	LPSPropValue	lpPropArray	= NULL;
	BOOL			bSuccess = false;

	SizedSPropTagArray(1, sptaBuffer) = {1,{ulPropTag}};
	*lppProp = NULL;

	WC_H(lpMAPIProp->GetProps((LPSPropTagArray)&sptaBuffer, 0, &cValues, &lpPropArray));

	if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag) && MAPI_E_NOT_ENOUGH_MEMORY == lpPropArray->Value.err)
	{
		MAPIFreeBuffer(lpPropArray);
		lpPropArray = NULL;
		//need to get the data as a stream
		LPSTREAM lpStream = NULL;

		WC_H(lpMAPIProp->OpenProperty(
			ulPropTag,
			&IID_IStream,
			STGM_READ,
			0,
			(LPUNKNOWN*) &lpStream));
		if (SUCCEEDED(hRes) && lpStream)
		{
			STATSTG	StatInfo = {0};
			lpStream->Stat(&StatInfo, STATFLAG_NONAME); //find out how much space we need

			// We're not going to try to support MASSIVE properties.
			if (!StatInfo.cbSize.HighPart)
			{
				EC_H(MAPIAllocateBuffer(
					sizeof(SPropValue),
					(LPVOID*) &lpPropArray));
				if (lpPropArray)
				{
					memset(lpPropArray,0,sizeof(SPropValue));
					lpPropArray->ulPropTag = ulPropTag;

					if (StatInfo.cbSize.LowPart)
					{
						EC_H(MAPIAllocateMore(
							StatInfo.cbSize.LowPart,
							lpPropArray,
							(LPVOID*) &lpPropArray->Value.bin.lpb));
						if (lpPropArray->Value.bin.lpb)
						{
							EC_H(lpStream->Read(lpPropArray->Value.bin.lpb, StatInfo.cbSize.LowPart, &lpPropArray->Value.bin.cb));
							if (SUCCEEDED(hRes) && lpPropArray->Value.bin.cb == StatInfo.cbSize.LowPart)
							{
								bSuccess = true;
							}
						}
					}
					else bSuccess = true; // if LowPart was NULL, we return the empty buffer
				}
			}
		}
		if (lpStream) lpStream->Release();
	}
	else if (lpPropArray && cValues == 1 && lpPropArray->ulPropTag == ulPropTag)
	{
		bSuccess = true;
	}

	if (bSuccess)
	{
		*lppProp = lpPropArray;
	}
	else
	{
		MAPIFreeBuffer(lpPropArray);
		if (SUCCEEDED(hRes)) hRes = MAPI_E_CALL_FAILED;
	}

	return hRes;
}

void InterpretBinaryAsString(SBinary myBin, MAPIStructType myStructType, LPMAPIPROP lpMAPIProp, ULONG ulPropTag, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	LPTSTR szResultString = NULL;

	switch (myStructType)
	{
	case stTimeZoneDefinition:
		TimeZoneDefinitionToString(myBin,&szResultString);
		break;
	case stTimeZone:
		TimeZoneToString(myBin,&szResultString);
		break;
	case stSecurityDescriptor:
		SDBinToString(myBin,lpMAPIProp,ulPropTag,&szResultString);
		break;
	case stExtendedFolderFlags:
		ExtendedFlagsBinToString(myBin,&szResultString);
		break;
	case stAppointmentRecurrencePattern:
		AppointmentRecurrencePatternToString(myBin,&szResultString);
		break;
	case stRecurrencePattern:
		RecurrencePatternToString(myBin,&szResultString);
		break;
	case stReportTag:
		ReportTagToString(myBin,&szResultString);
		break;
	case stConversationIndex:
		ConversationIndexToString(myBin,&szResultString);
		break;
	case stTaskAssigners:
		TaskAssignersToString(myBin,&szResultString);
		break;
	case stGlobalObjectId:
		GlobalObjectIdToString(myBin,&szResultString);
		break;
	case stOneOffEntryId:
		OneOffEntryIdToString(myBin,&szResultString);
	}
	if (szResultString) *lpszResultString = szResultString;
}

class CBinaryParser
{
public:
	CBinaryParser(size_t cbBin, LPBYTE lpBin);

	void Advance(size_t cbAdvance);
	size_t GetCurrentOffset();
	size_t RemainingBytes();
	void GetBYTE(BYTE* pBYTE);
	void GetWORD(WORD* pWORD);
	void GetDWORD(DWORD* pDWORD);
	void GetLARGE_INTEGER(LARGE_INTEGER* pLARGE_INTEGER);
	void GetBYTES(size_t cbBytes, LPBYTE* ppBYTES);
	void GetBYTESNoAlloc(size_t cbBytes, LPBYTE pBYTES);
	void GetStringA(size_t cchChar, LPSTR* ppStr);
	void GetStringW(size_t cchChar, LPWSTR* ppStr);
	void GetStringA(LPSTR* ppStr);
	void GetStringW(LPWSTR* ppStr);

private:
	size_t m_cbBin;
	LPBYTE m_lpBin;
	LPBYTE m_lpEnd;
	LPBYTE m_lpCur;
};

CBinaryParser::CBinaryParser(size_t cbBin, LPBYTE lpBin)
{
	m_cbBin = cbBin;
	m_lpBin = lpBin;
	m_lpCur = lpBin;
	m_lpEnd = lpBin+cbBin;
}

void CBinaryParser::Advance(size_t cbAdvance)
{
	m_lpCur += cbAdvance;
}

size_t CBinaryParser::GetCurrentOffset()
{
	return m_lpCur - m_lpBin;
}

// If we're before the end of the buffer, return the count of remaining bytes
// If we're at or past the end of the buffer, return 0;
size_t CBinaryParser::RemainingBytes()
{
	if (m_lpCur < m_lpEnd) return m_lpEnd - m_lpCur;
	return 0;
}


void CBinaryParser::GetBYTE(BYTE* pBYTE)
{
	if (!pBYTE || !m_lpCur) return;
	if (m_lpCur + sizeof(BYTE) > m_lpEnd) return;
	*pBYTE = *((BYTE*)m_lpCur);
	m_lpCur += sizeof(BYTE);
}

void CBinaryParser::GetWORD(WORD* pWORD)
{
	if (!pWORD || !m_lpCur) return;
	if (m_lpCur + sizeof(WORD) > m_lpEnd) return;
	*pWORD = *((WORD*)m_lpCur);
	m_lpCur += sizeof(WORD);
}

void CBinaryParser::GetDWORD(DWORD* pDWORD)
{
	if (!pDWORD || !m_lpCur) return;
	if (m_lpCur + sizeof(DWORD) > m_lpEnd) return;
	*pDWORD = *((DWORD*)m_lpCur);
	m_lpCur += sizeof(DWORD);
}

void CBinaryParser::GetLARGE_INTEGER(LARGE_INTEGER* pLARGE_INTEGER)
{
	if (!pLARGE_INTEGER || !m_lpCur) return;
	if (m_lpCur + sizeof(LARGE_INTEGER) > m_lpEnd) return;
	*pLARGE_INTEGER = *((LARGE_INTEGER*)m_lpCur);
	m_lpCur += sizeof(LARGE_INTEGER);
}

void CBinaryParser::GetBYTES(size_t cbBytes, LPBYTE* ppBYTES)
{
	if (!cbBytes || !ppBYTES) return;
	if (m_lpCur + cbBytes > m_lpEnd) return;
	*ppBYTES = new BYTE[cbBytes];
	if (*ppBYTES)
	{
		memset(*ppBYTES,0,sizeof(BYTE) * cbBytes);
		memcpy(*ppBYTES,m_lpCur,cbBytes);
	}
	m_lpCur += cbBytes;
}

void CBinaryParser::GetBYTESNoAlloc(size_t cbBytes, LPBYTE pBYTES)
{
	if (!cbBytes || !pBYTES) return;
	if (m_lpCur + cbBytes > m_lpEnd) return;
	memset(pBYTES,0,sizeof(BYTE) * cbBytes);
	memcpy(pBYTES,m_lpCur,cbBytes);
	m_lpCur += cbBytes;
}

// cchChar is the length of the source string, NOT counting the NULL terminator
// If the NULL terminator is known to be in the stream, follow this call with a call to Advance(sizeof(CHAR))
void CBinaryParser::GetStringA(size_t cchChar, LPSTR* ppStr)
{
	if (!cchChar || !ppStr) return;
	if (m_lpCur + sizeof(CHAR) * cchChar > m_lpEnd) return;
	*ppStr = new CHAR[cchChar+1];
	if (*ppStr)
	{
		memset(*ppStr,0,sizeof(CHAR) * cchChar);
		memcpy(*ppStr,m_lpCur,sizeof(CHAR) * cchChar);
		(*ppStr)[cchChar] = NULL;
	}
	m_lpCur += sizeof(CHAR) * cchChar;
}

// cchChar is the length of the source string, NOT counting the NULL terminator
// If the NULL terminator is known to be in the stream, follow this call with a call to Advance(sizeof(WCHAR))
void CBinaryParser::GetStringW(size_t cchWChar, LPWSTR* ppStr)
{
	if (!cchWChar || !ppStr) return;
	if (m_lpCur + sizeof(WCHAR) * cchWChar > m_lpEnd) return;
	*ppStr = new WCHAR[cchWChar+1];
	if (*ppStr)
	{
		memset(*ppStr,0,sizeof(WCHAR) * cchWChar);
		memcpy(*ppStr,m_lpCur,sizeof(WCHAR) * cchWChar);
		(*ppStr)[cchWChar] = NULL;
	}
	m_lpCur += sizeof(WCHAR) * cchWChar;
}

// No size specified - assume the NULL terminator is in the stream, but don't read off the end
void CBinaryParser::GetStringA(LPSTR* ppStr)
{
	if (!ppStr) return;
	size_t cchChar = NULL;
	HRESULT hRes = S_OK;

	hRes = StringCchLengthA((LPSTR)m_lpCur,(m_lpEnd-m_lpCur)/sizeof(CHAR),&cchChar);

	if (FAILED(hRes)) return;

	// With string length in hand, we defer to our other implementation
	// Add 1 for the NULL terminator
	return GetStringA(cchChar+1,ppStr);
}

// No size specified - assume the NULL terminator is in the stream, but don't read off the end
void CBinaryParser::GetStringW(LPWSTR* ppStr)
{
	if (!ppStr) return;
	size_t cchChar = NULL;
	HRESULT hRes = S_OK;

	hRes = StringCchLengthW((LPWSTR)m_lpCur,(m_lpEnd-m_lpCur)/sizeof(WCHAR),&cchChar);

	if (FAILED(hRes)) return;

	// With string length in hand, we defer to our other implementation
	// Add 1 for the NULL terminator
	return GetStringW(cchChar+1,ppStr);
}

CString JunkDataToString(size_t cbJunkData, LPBYTE lpJunkData)
{
	if (!cbJunkData || !lpJunkData) return _T(""); // STRING_OK
	CString szTmp;
	SBinary sBin = {0};

	sBin.cb = (ULONG) cbJunkData;
	sBin.lpb = lpJunkData;
	szTmp.FormatMessage(IDS_JUNKDATASIZE,
		cbJunkData,
		BinToHexString(&sBin,true));
	return szTmp;
}

// result allocated with new
// clean up with delete[]
LPTSTR CStringToString(CString szCString)
{
	size_t cchCString = szCString.GetLength()+1;
	LPTSTR szOut = new TCHAR[cchCString];
	if (szOut)
	{
		HRESULT hRes = S_OK;
		EC_H(StringCchCopy(szOut,cchCString,(LPCTSTR)szCString));
	}
	return szOut;
}

void AppointmentRecurrencePatternToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	AppointmentRecurrencePatternStruct* parpPattern = BinToAppointmentRecurrencePatternStruct(myBin.cb,myBin.lpb);
	if (parpPattern)
	{
		*lpszResultString = AppointmentRecurrencePatternStructToString(parpPattern);
		DeleteAppointmentRecurrencePatternStruct(parpPattern);
	}
}

void RecurrencePatternToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	RecurrencePatternStruct* prpPattern = BinToRecurrencePatternStruct(myBin.cb,myBin.lpb,NULL);
	if (prpPattern)
	{
		*lpszResultString = RecurrencePatternStructToString(prpPattern);
		DeleteRecurrencePatternStruct(prpPattern);
	}
}

void DeleteAppointmentRecurrencePatternStruct(AppointmentRecurrencePatternStruct* parpPattern)
{
	if (!parpPattern) return;
	delete[] parpPattern->RecurrencePattern;
	int i = 0;
	if (parpPattern->ExceptionCount && parpPattern->ExceptionInfo)
	{
		for (i = 0; i < parpPattern->ExceptionCount ; i++)
		{
			delete[] parpPattern->ExceptionInfo[i].Subject;
			delete[] parpPattern->ExceptionInfo[i].Location;
		}
	}
	delete[] parpPattern->ExceptionInfo;
	if (parpPattern->ExceptionCount && parpPattern->ExtendedException)
	{
		for (i = 0; i < parpPattern->ExceptionCount ; i++)
		{
			delete[] parpPattern->ExtendedException[i].ChangeHighlight.Reserved;
			delete[] parpPattern->ExtendedException[i].ReservedBlockEE1;
			delete[] parpPattern->ExtendedException[i].WideCharSubject;
			delete[] parpPattern->ExtendedException[i].WideCharLocation;
			delete[] parpPattern->ExtendedException[i].ReservedBlockEE2;
		}
	}
	delete[] parpPattern->ExtendedException;
	delete[] parpPattern->ReservedBlock2;
	delete[] parpPattern->JunkData;
	delete parpPattern;
}

void DeleteRecurrencePatternStruct(RecurrencePatternStruct* prpPattern)
{
	if (!prpPattern) return;
	delete[] prpPattern->DeletedInstanceDates;
	delete[] prpPattern->ModifiedInstanceDates;
	delete prpPattern;
}

// There may be recurrence with over 500 exceptions, but we're not going to try to parse them
#define _MaxExceptions 500
#define _MaxReservedBlock 0xffff

// Allocates return value with new.
// Clean up with DeleteAppointmentRecurrencePatternStruct.
AppointmentRecurrencePatternStruct* BinToAppointmentRecurrencePatternStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	AppointmentRecurrencePatternStruct arpPattern = {0};
	CBinaryParser Parser(cbBin,lpBin);

	size_t cbBinRead = 0;
	arpPattern.RecurrencePattern = BinToRecurrencePatternStruct(cbBin,lpBin,&cbBinRead);
	Parser.Advance(cbBinRead);
	Parser.GetDWORD(&arpPattern.ReaderVersion2);
	Parser.GetDWORD(&arpPattern.WriterVersion2);
	Parser.GetDWORD(&arpPattern.StartTimeOffset);
	Parser.GetDWORD(&arpPattern.EndTimeOffset);
	Parser.GetWORD(&arpPattern.ExceptionCount);

	if (arpPattern.ExceptionCount &&
		arpPattern.ExceptionCount == arpPattern.RecurrencePattern->ModifiedInstanceCount &&
		arpPattern.ExceptionCount < _MaxExceptions)
	{
		arpPattern.ExceptionInfo = new ExceptionInfoStruct[arpPattern.ExceptionCount];
		if (arpPattern.ExceptionInfo)
		{
			memset(arpPattern.ExceptionInfo,0,sizeof(ExceptionInfoStruct) * arpPattern.ExceptionCount);
			WORD i = 0;
			for (i = 0 ; i < arpPattern.ExceptionCount ; i++)
			{
				Parser.GetDWORD(&arpPattern.ExceptionInfo[i].StartDateTime);
				Parser.GetDWORD(&arpPattern.ExceptionInfo[i].EndDateTime);
				Parser.GetDWORD(&arpPattern.ExceptionInfo[i].OriginalStartDate);
				Parser.GetWORD(&arpPattern.ExceptionInfo[i].OverrideFlags);
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
				{
					Parser.GetWORD(&arpPattern.ExceptionInfo[i].SubjectLength);
					Parser.GetWORD(&arpPattern.ExceptionInfo[i].SubjectLength2);
					if (arpPattern.ExceptionInfo[i].SubjectLength2 && arpPattern.ExceptionInfo[i].SubjectLength2 + 1 == arpPattern.ExceptionInfo[i].SubjectLength)
					{
						Parser.GetStringA(arpPattern.ExceptionInfo[i].SubjectLength2,&arpPattern.ExceptionInfo[i].Subject);
					}
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_MEETINGTYPE)
				{
					Parser.GetDWORD(&arpPattern.ExceptionInfo[i].MeetingType);
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_REMINDERDELTA)
				{
					Parser.GetDWORD(&arpPattern.ExceptionInfo[i].ReminderDelta);
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_REMINDER)
				{
					Parser.GetDWORD(&arpPattern.ExceptionInfo[i].ReminderSet);
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					Parser.GetWORD(&arpPattern.ExceptionInfo[i].LocationLength);
					Parser.GetWORD(&arpPattern.ExceptionInfo[i].LocationLength2);
					if (arpPattern.ExceptionInfo[i].LocationLength2 && arpPattern.ExceptionInfo[i].LocationLength2 + 1 == arpPattern.ExceptionInfo[i].LocationLength)
					{
						Parser.GetStringA(arpPattern.ExceptionInfo[i].LocationLength2,&arpPattern.ExceptionInfo[i].Location);
					}
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_BUSYSTATUS)
				{
					Parser.GetDWORD(&arpPattern.ExceptionInfo[i].BusyStatus);
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_ATTACHMENT)
				{
					Parser.GetDWORD(&arpPattern.ExceptionInfo[i].Attachment);
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_SUBTYPE)
				{
					Parser.GetDWORD(&arpPattern.ExceptionInfo[i].SubType);
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_APPTCOLOR)
				{
					Parser.GetDWORD(&arpPattern.ExceptionInfo[i].AppointmentColor);
				}
			}
		}
	}
	Parser.GetDWORD(&arpPattern.ReservedBlock1Size);
	if (arpPattern.ReservedBlock1Size && arpPattern.ReservedBlock1Size < _MaxReservedBlock)
	{
		Parser.GetBYTES(arpPattern.ReservedBlock1Size,&arpPattern.ReservedBlock1);
	}
	if (arpPattern.ExceptionCount &&
		arpPattern.ExceptionCount == arpPattern.RecurrencePattern->ModifiedInstanceCount &&
		arpPattern.ExceptionCount < _MaxExceptions &&
		arpPattern.ExceptionInfo)
	{
		arpPattern.ExtendedException = new ExtendedExceptionStruct[arpPattern.ExceptionCount];
		if (arpPattern.ExtendedException)
		{
			memset(arpPattern.ExtendedException,0,sizeof(ExtendedExceptionStruct) * arpPattern.ExceptionCount);
			WORD i = 0;
			for (i = 0 ; i < arpPattern.ExceptionCount ; i++)
			{
				if (arpPattern.WriterVersion2 >= 0x0003009)
				{
					Parser.GetDWORD(&arpPattern.ExtendedException[i].ChangeHighlight.ChangeHighlightSize);
					Parser.GetDWORD(&arpPattern.ExtendedException[i].ChangeHighlight.ChangeHighlightValue);
					if (arpPattern.ExtendedException[i].ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
					{
						Parser.GetBYTES(arpPattern.ExtendedException[i].ChangeHighlight.ChangeHighlightSize - sizeof(DWORD),&arpPattern.ExtendedException[i].ChangeHighlight.Reserved);
					}
				}
				Parser.GetDWORD(&arpPattern.ExtendedException[i].ReservedBlockEE1Size);
				if (arpPattern.ExtendedException[i].ReservedBlockEE1Size && arpPattern.ExtendedException[i].ReservedBlockEE1Size < _MaxReservedBlock)
				{
					Parser.GetBYTES(arpPattern.ExtendedException[i].ReservedBlockEE1Size,&arpPattern.ExtendedException[i].ReservedBlockEE1);
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
					arpPattern.ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					Parser.GetDWORD(&arpPattern.ExtendedException[i].StartDateTime);
					Parser.GetDWORD(&arpPattern.ExtendedException[i].EndDateTime);
					Parser.GetDWORD(&arpPattern.ExtendedException[i].OriginalStartDate);
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
				{
					Parser.GetWORD(&arpPattern.ExtendedException[i].WideCharSubjectLength);
					if (arpPattern.ExtendedException[i].WideCharSubjectLength)
					{
						Parser.GetStringW(arpPattern.ExtendedException[i].WideCharSubjectLength,&arpPattern.ExtendedException[i].WideCharSubject);
					}
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					Parser.GetWORD(&arpPattern.ExtendedException[i].WideCharLocationLength);
					if (arpPattern.ExtendedException[i].WideCharLocationLength)
					{
						Parser.GetStringW(arpPattern.ExtendedException[i].WideCharLocationLength,&arpPattern.ExtendedException[i].WideCharLocation);
					}
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
					arpPattern.ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					Parser.GetDWORD(&arpPattern.ExtendedException[i].ReservedBlockEE2Size);
					if (arpPattern.ExtendedException[i].ReservedBlockEE2Size && arpPattern.ExtendedException[i].ReservedBlockEE2Size < _MaxReservedBlock)
					{
						Parser.GetBYTES(arpPattern.ExtendedException[i].ReservedBlockEE2Size,&arpPattern.ExtendedException[i].ReservedBlockEE2);
					}
				}
			}
		}
	}
	Parser.GetDWORD(&arpPattern.ReservedBlock2Size);
	if (arpPattern.ReservedBlock2Size && arpPattern.ReservedBlock2Size < _MaxReservedBlock)
	{
		Parser.GetBYTES(arpPattern.ReservedBlock2Size,&arpPattern.ReservedBlock2);
	}

	// Junk data remains
	if (Parser.RemainingBytes() > 0)
	{
		arpPattern.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(arpPattern.JunkDataSize,&arpPattern.JunkData);
	}

	AppointmentRecurrencePatternStruct* parpPattern = new AppointmentRecurrencePatternStruct;
	if (parpPattern)
	{
		*parpPattern = arpPattern;
	}

	return parpPattern;
}

// Parses lpBin as a RecurrencePatternStruct
// lpcbBytesRead returns the number of bytes consumed
// Allocates return value with new.
// clean up with delete.
RecurrencePatternStruct* BinToRecurrencePatternStruct(ULONG cbBin, LPBYTE lpBin, size_t* lpcbBytesRead)
{
	if (!lpBin) return NULL;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	CBinaryParser Parser(cbBin,lpBin);

	RecurrencePatternStruct rpPattern = {0};

	Parser.GetWORD(&rpPattern.ReaderVersion);
	Parser.GetWORD(&rpPattern.WriterVersion);
	Parser.GetWORD(&rpPattern.RecurFrequency);
	Parser.GetWORD(&rpPattern.PatternType);
	Parser.GetWORD(&rpPattern.CalendarType);
	Parser.GetDWORD(&rpPattern.FirstDateTime);
	Parser.GetDWORD(&rpPattern.Period);
	Parser.GetDWORD(&rpPattern.SlidingFlag);

	switch (rpPattern.PatternType)
	{
	case rptMinute:
		break;
	case rptWeek:
		Parser.GetDWORD(&rpPattern.PatternTypeSpecific.WeekRecurrencePattern);
		break;
	case rptMonth:
	case rptMonthEnd:
	case rptHjMonth:
	case rptHjMonthEnd:
		Parser.GetDWORD(&rpPattern.PatternTypeSpecific.MonthRecurrencePattern);
		break;
	case rptMonthNth:
	case rptHjMonthNth:
		Parser.GetDWORD(&rpPattern.PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek);
		Parser.GetDWORD(&rpPattern.PatternTypeSpecific.MonthNthRecurrencePattern.N);
		break;
	}

	Parser.GetDWORD(&rpPattern.EndType);
	Parser.GetDWORD(&rpPattern.OccurrenceCount);
	Parser.GetDWORD(&rpPattern.FirstDOW);
	Parser.GetDWORD(&rpPattern.DeletedInstanceCount);

	if (rpPattern.DeletedInstanceCount && rpPattern.DeletedInstanceCount < _MaxExceptions)
	{
		rpPattern.DeletedInstanceDates = new DWORD[rpPattern.DeletedInstanceCount];
		if (rpPattern.DeletedInstanceDates)
		{
			memset(rpPattern.DeletedInstanceDates,0,sizeof(DWORD) * rpPattern.DeletedInstanceCount);
			DWORD i = 0;
			for (i = 0 ; i < rpPattern.DeletedInstanceCount ; i++)
			{
				Parser.GetDWORD(&rpPattern.DeletedInstanceDates[i]);
			}
		}
	}

	Parser.GetDWORD(&rpPattern.ModifiedInstanceCount);

	if (rpPattern.ModifiedInstanceCount &&
		rpPattern.ModifiedInstanceCount <= rpPattern.DeletedInstanceCount &&
		rpPattern.ModifiedInstanceCount < _MaxExceptions)
	{
		rpPattern.ModifiedInstanceDates = new DWORD[rpPattern.ModifiedInstanceCount];
		if (rpPattern.ModifiedInstanceDates)
		{
			memset(rpPattern.ModifiedInstanceDates,0,sizeof(DWORD) * rpPattern.ModifiedInstanceCount);
			DWORD i = 0;
			for (i = 0 ; i < rpPattern.ModifiedInstanceCount ; i++)
			{
				Parser.GetDWORD(&rpPattern.ModifiedInstanceDates[i]);
			}
		}
	}
	Parser.GetDWORD(&rpPattern.StartDate);
	Parser.GetDWORD(&rpPattern.EndDate);

	// Junk data remains
	// Only fill out junk data if we've not been asked to report back how many bytes we read
	// If we've been asked to report back, then someone else will handle the remaining data
	if (!lpcbBytesRead && Parser.RemainingBytes() > 0)
	{
		rpPattern.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(rpPattern.JunkDataSize,&rpPattern.JunkData);
	}

	RecurrencePatternStruct* prpPattern = new RecurrencePatternStruct;
	if (prpPattern)
	{
		*prpPattern = rpPattern;
		if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();
	}

	return prpPattern;
}

// result allocated with new
// clean up with delete[]
LPTSTR AppointmentRecurrencePatternStructToString(AppointmentRecurrencePatternStruct* parpPattern)
{
	if (!parpPattern) return NULL;

	CString szARP;
	CString szTmp;
	HRESULT hRes = S_OK;
	szARP = RecurrencePatternStructToString(parpPattern->RecurrencePattern);

	szTmp.FormatMessage(IDS_ARPHEADER,
		parpPattern->ReaderVersion2,
		parpPattern->WriterVersion2,
		parpPattern->StartTimeOffset,RTimeToString(parpPattern->StartTimeOffset),
		parpPattern->EndTimeOffset,RTimeToString(parpPattern->EndTimeOffset),
		parpPattern->ExceptionCount);
	szARP += szTmp;

	WORD i = 0;
	if (parpPattern->ExceptionCount && parpPattern->ExceptionInfo)
	{
		for (i = 0; i < parpPattern->ExceptionCount ; i++)
		{
			CString szExceptionInfo;
			LPTSTR szOverrideFlags = NULL;
			EC_H(InterpretFlags(flagOverrideFlags, parpPattern->ExceptionInfo[i].OverrideFlags, &szOverrideFlags));
			szExceptionInfo.FormatMessage(IDS_ARPEXHEADER,
				i,parpPattern->ExceptionInfo[i].StartDateTime,RTimeToString(parpPattern->ExceptionInfo[i].StartDateTime),
				parpPattern->ExceptionInfo[i].EndDateTime,RTimeToString(parpPattern->ExceptionInfo[i].EndDateTime),
				parpPattern->ExceptionInfo[i].OriginalStartDate,RTimeToString(parpPattern->ExceptionInfo[i].OriginalStartDate),
				parpPattern->ExceptionInfo[i].OverrideFlags,szOverrideFlags);
			delete[] szOverrideFlags;
			szOverrideFlags = NULL;
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
			{
				szTmp.FormatMessage(IDS_ARPEXSUBJECT,
					i,parpPattern->ExceptionInfo[i].SubjectLength,
					parpPattern->ExceptionInfo[i].SubjectLength2,
					parpPattern->ExceptionInfo[i].Subject?parpPattern->ExceptionInfo[i].Subject:"");
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_MEETINGTYPE)
			{
				LPTSTR szFlags = NULL;
				EC_H(InterpretFlags(PROP_TAG((guidPSETID_Appointment),dispidApptStateFlags), parpPattern->ExceptionInfo[i].MeetingType, &szFlags));
				szTmp.FormatMessage(IDS_ARPEXMEETINGTYPE,
					i,parpPattern->ExceptionInfo[i].MeetingType,szFlags);
				delete[] szFlags;
				szFlags = NULL;
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_REMINDERDELTA)
			{
				szTmp.FormatMessage(IDS_ARPEXREMINDERDELTA,
					i,parpPattern->ExceptionInfo[i].ReminderDelta);
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_REMINDER)
			{
				szTmp.FormatMessage(IDS_ARPEXREMINDERSET,
					i,parpPattern->ExceptionInfo[i].ReminderSet);
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				szTmp.FormatMessage(IDS_ARPEXLOCATION,
					i,parpPattern->ExceptionInfo[i].LocationLength,
					parpPattern->ExceptionInfo[i].LocationLength2,
					parpPattern->ExceptionInfo[i].Location?parpPattern->ExceptionInfo[i].Location:"");
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_BUSYSTATUS)
			{
				LPTSTR szFlags = NULL;
				EC_H(InterpretFlags(PROP_TAG((guidPSETID_Common),dispidBusyStatus), parpPattern->ExceptionInfo[i].BusyStatus, &szFlags));
				szTmp.FormatMessage(IDS_ARPEXBUSYSTATUS,
					i,parpPattern->ExceptionInfo[i].BusyStatus,szFlags);
				delete[] szFlags;
				szFlags = NULL;
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_ATTACHMENT)
			{
				szTmp.FormatMessage(IDS_ARPEXATTACHMENT,
					i,parpPattern->ExceptionInfo[i].Attachment);
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_SUBTYPE)
			{
				szTmp.FormatMessage(IDS_ARPEXSUBTYPE,
					i,parpPattern->ExceptionInfo[i].SubType);
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_APPTCOLOR)
			{
				szTmp.FormatMessage(IDS_ARPEXAPPOINTMENTCOLOR,
					i,parpPattern->ExceptionInfo[i].AppointmentColor);
				szExceptionInfo += szTmp;
			}
			szARP += szExceptionInfo;
		}
	}

	szTmp.FormatMessage(IDS_ARPRESERVED1,
		parpPattern->ReservedBlock1Size);
	szARP += szTmp;
	if (parpPattern->ReservedBlock1Size)
	{
		SBinary sBin = {0};
		sBin.cb = parpPattern->ReservedBlock1Size;
		sBin.lpb = parpPattern->ReservedBlock1;
		szARP += BinToHexString(&sBin,true);
	}

	if (parpPattern->ExceptionCount && parpPattern->ExtendedException)
	{
		for (i = 0; i < parpPattern->ExceptionCount ; i++)
		{
			CString szExtendedException;
			if (parpPattern->WriterVersion2 >= 0x00003009)
			{
				LPTSTR szFlags = NULL;
				EC_H(InterpretFlags(PROP_TAG((guidPSETID_Appointment),dispidChangeHighlight), parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightValue, &szFlags));
				szTmp.FormatMessage(IDS_ARPEXCHANGEHIGHLIGHT,
					i,parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightSize,
					parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightValue,szFlags);
				delete[] szFlags;
				szFlags = NULL;
				szExtendedException += szTmp;
				if (parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
				{
					szTmp.FormatMessage(IDS_ARPEXCHANGEHIGHLIGHTRESERVED,
						i);
					szExtendedException += szTmp;
					SBinary sBin = {0};
					sBin.cb = parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightSize - sizeof(DWORD);
					sBin.lpb = parpPattern->ExtendedException[i].ChangeHighlight.Reserved;
					szExtendedException += BinToHexString(&sBin,true);
					szExtendedException += _T("\n"); // STRING_OK
				}
			}
			szTmp.FormatMessage(IDS_ARPEXRESERVED1,
				i,parpPattern->ExtendedException[i].ReservedBlockEE1Size);
			szExtendedException += szTmp;
			if (parpPattern->ExtendedException[i].ReservedBlockEE1Size)
			{
				SBinary sBin = {0};
				sBin.cb = parpPattern->ExtendedException[i].ReservedBlockEE1Size;
				sBin.lpb = parpPattern->ExtendedException[i].ReservedBlockEE1;
				szExtendedException += BinToHexString(&sBin,true);
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
				parpPattern->ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				szTmp.FormatMessage(IDS_ARPEXDATETIME,
					i,parpPattern->ExtendedException[i].StartDateTime,RTimeToString(parpPattern->ExtendedException[i].StartDateTime),
					parpPattern->ExtendedException[i].EndDateTime,RTimeToString(parpPattern->ExtendedException[i].EndDateTime),
					parpPattern->ExtendedException[i].OriginalStartDate,RTimeToString(parpPattern->ExtendedException[i].OriginalStartDate));
				szExtendedException += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
			{
				szTmp.FormatMessage(IDS_ARPEXWIDESUBJECT,
					i,parpPattern->ExtendedException[i].WideCharSubjectLength,
					parpPattern->ExtendedException[i].WideCharSubject?parpPattern->ExtendedException[i].WideCharSubject:L"");
				szExtendedException += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				szTmp.FormatMessage(IDS_ARPEXWIDELOCATION,
					i,parpPattern->ExtendedException[i].WideCharLocationLength,
					parpPattern->ExtendedException[i].WideCharLocation?parpPattern->ExtendedException[i].WideCharLocation:L"");
				szExtendedException += szTmp;
			}
			szTmp.FormatMessage(IDS_ARPEXRESERVED1,
				i,parpPattern->ExtendedException[i].ReservedBlockEE2Size);
			szExtendedException += szTmp;
			if (parpPattern->ExtendedException[i].ReservedBlockEE2Size)
			{
				SBinary sBin = {0};
				sBin.cb = parpPattern->ExtendedException[i].ReservedBlockEE2Size;
				sBin.lpb = parpPattern->ExtendedException[i].ReservedBlockEE2;
				szExtendedException += BinToHexString(&sBin,true);
			}

			szARP += szExtendedException;
		}
	}

	szTmp.FormatMessage(IDS_ARPRESERVED2,
		parpPattern->ReservedBlock2Size);
	szARP += szTmp;
	if (parpPattern->ReservedBlock2Size)
	{
		SBinary sBin = {0};
		sBin.cb = parpPattern->ReservedBlock2Size;
		sBin.lpb = parpPattern->ReservedBlock2;
		szARP += BinToHexString(&sBin,true);
	}

	szARP += JunkDataToString(parpPattern->JunkDataSize,parpPattern->JunkData);

	return CStringToString(szARP);
} // AppointmentRecurrencePatternStructToString

// result allocated with new
// clean up with delete[]
LPTSTR RecurrencePatternStructToString(RecurrencePatternStruct* prpPattern)
{
	if (!prpPattern) return NULL;

	CString szRP;
	CString szTmp;
	HRESULT hRes = S_OK;

	LPTSTR szRecurFrequency = NULL;
	LPTSTR szPatternType = NULL;
	LPTSTR szCalendarType = NULL;
	EC_H(InterpretFlags(flagRecurFrequency, prpPattern->RecurFrequency, &szRecurFrequency));
	EC_H(InterpretFlags(flagPatternType, prpPattern->PatternType, &szPatternType));
	EC_H(InterpretFlags(flagCalendarType, prpPattern->CalendarType, &szCalendarType));
	szTmp.FormatMessage(IDS_RPHEADER,
		prpPattern->ReaderVersion,
		prpPattern->WriterVersion,
		prpPattern->RecurFrequency,szRecurFrequency,
		prpPattern->PatternType,szPatternType,
		prpPattern->CalendarType,szCalendarType,
		prpPattern->FirstDateTime,
		prpPattern->Period,
		prpPattern->SlidingFlag);
	delete[] szRecurFrequency;
	delete[] szPatternType;
	delete[] szCalendarType;
	szRecurFrequency = NULL;
	szPatternType = NULL;
	szCalendarType = NULL;
	szRP += szTmp;

	LPTSTR szDOW = NULL;
	LPTSTR szN = NULL;
	switch (prpPattern->PatternType)
	{
	case rptMinute:
		break;
	case rptWeek:
		EC_H(InterpretFlags(flagDOW, prpPattern->PatternTypeSpecific.WeekRecurrencePattern, &szDOW));
		szTmp.FormatMessage(IDS_RPPATTERNWEEK,
			prpPattern->PatternTypeSpecific.WeekRecurrencePattern,szDOW);
		szRP += szTmp;
		break;
	case rptMonth:
	case rptMonthEnd:
	case rptHjMonth:
	case rptHjMonthEnd:
		szTmp.FormatMessage(IDS_RPPATTERNMONTH,
			prpPattern->PatternTypeSpecific.MonthRecurrencePattern);
		szRP += szTmp;
		break;
	case rptMonthNth:
	case rptHjMonthNth:
		EC_H(InterpretFlags(flagDOW, prpPattern->PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek, &szDOW));
		EC_H(InterpretFlags(flagN, prpPattern->PatternTypeSpecific.MonthNthRecurrencePattern.N, &szN));
		szTmp.FormatMessage(IDS_RPPATTERNMONTHNTH,
			prpPattern->PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek, szDOW,
			prpPattern->PatternTypeSpecific.MonthNthRecurrencePattern.N,szN);
		szRP += szTmp;
		break;
	}
	delete[] szDOW;
	delete[] szN;
	szDOW = NULL;
	szN = NULL;

	LPTSTR szEndType = NULL;
	LPTSTR szFirstDOW = NULL;
	EC_H(InterpretFlags(flagEndType, prpPattern->EndType, &szEndType));
	EC_H(InterpretFlags(flagFirstDOW, prpPattern->FirstDOW, &szFirstDOW));

	szTmp.FormatMessage(IDS_RPHEADER2,
		prpPattern->EndType, szEndType,
		prpPattern->OccurrenceCount,
		prpPattern->FirstDOW,szFirstDOW,
		prpPattern->DeletedInstanceCount);
	szRP += szTmp;
	delete[] szEndType;
	delete[] szFirstDOW;
	szEndType = NULL;
	szFirstDOW = NULL;

	if (prpPattern->DeletedInstanceCount && prpPattern->DeletedInstanceDates)
	{
		DWORD i = 0;
		for (i = 0 ; i < prpPattern->DeletedInstanceCount ; i++)
		{
			szTmp.FormatMessage(IDS_RPDELETEDINSTANCEDATES,
				i,prpPattern->DeletedInstanceDates[i],RTimeToString(prpPattern->DeletedInstanceDates[i]));
			szRP += szTmp;
		}
	}

	szTmp.FormatMessage(IDS_RPMODIFIEDINSTANCECOUNT,
		prpPattern->ModifiedInstanceCount);
	szRP += szTmp;

	if (prpPattern->ModifiedInstanceCount && prpPattern->ModifiedInstanceDates)
	{
		DWORD i = 0;
		for (i = 0 ; i < prpPattern->ModifiedInstanceCount ; i++)
		{
			szTmp.FormatMessage(IDS_RPMODIFIEDINSTANCEDATES,
				i,prpPattern->ModifiedInstanceDates[i],RTimeToString(prpPattern->ModifiedInstanceDates[i]));
			szRP += szTmp;
		}
	}

	szTmp.FormatMessage(IDS_RPDATE,
		prpPattern->StartDate,RTimeToString(prpPattern->StartDate),
		prpPattern->EndDate,RTimeToString(prpPattern->EndDate));
	szRP += szTmp;

	szRP += JunkDataToString(prpPattern->JunkDataSize,prpPattern->JunkData);

	return CStringToString(szRP);
} // RecurrencePatternStructToString

void SDBinToString(SBinary myBin, LPMAPIPROP lpMAPIProp, ULONG ulPropTag, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	HRESULT hRes = S_OK;
	LPBYTE lpSDToParse = myBin.lpb;

	if (lpSDToParse)
	{
		eAceType acetype = acetypeMessage;
		switch (GetMAPIObjectType(lpMAPIProp))
		{
		case (MAPI_STORE):
		case (MAPI_ADDRBOOK):
		case (MAPI_FOLDER):
		case (MAPI_ABCONT):
			acetype = acetypeContainer;
			break;
		}

		if (PR_FREEBUSY_NT_SECURITY_DESCRIPTOR == ulPropTag)
			acetype = acetypeFreeBusy;

		CString szDACL;
		CString szInfo;

		EC_H(SDToString(lpSDToParse, acetype, &szDACL, &szInfo));

		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagSecurityVersion, SECURITY_DESCRIPTOR_VERSION(lpSDToParse), &szFlags));

		CString szResult;
		szResult.FormatMessage(IDS_SECURITYDESCRIPTOR,(LPCTSTR) szInfo,SECURITY_DESCRIPTOR_VERSION(lpSDToParse),szFlags,(LPCTSTR) szDACL);

		delete[] szFlags;
		szFlags = NULL;

		size_t cchSD = szResult.GetLength()+1;
		*lpszResultString = new TCHAR[cchSD];
		if (*lpszResultString)
		{
			EC_H(StringCchCopy(*lpszResultString,cchSD,(LPCTSTR)szResult));
		}
	}
}

// Allocates lpszResultString with new, free with delete[]
void ExtendedFlagsBinToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	ExtendedFlagsStruct* pefExtendedFlags = BinToExtendedFlagsStruct(myBin.cb,myBin.lpb);
	if (pefExtendedFlags)
	{
		*lpszResultString = ExtendedFlagsStructToString(pefExtendedFlags);
		DeleteExtendedFlagsStruct(pefExtendedFlags);
	}
}

ExtendedFlagsStruct* BinToExtendedFlagsStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	ExtendedFlagsStruct efExtendedFlags = {0};
	CBinaryParser Parser(cbBin,lpBin);

	// Run through the parser once to count the number of flag structs
	for (;;)
	{
		// Must have at least 2 bytes left to have another flag
		if (Parser.RemainingBytes() <  2) break;
		BYTE ulId = NULL;
		BYTE cbData = NULL;
		Parser.GetBYTE(&ulId);
		Parser.GetBYTE(&cbData);
		// Must have at least cbData bytes left to be a valid flag
		if (Parser.RemainingBytes() < cbData) break;

		Parser.Advance(cbData);
		efExtendedFlags.ulNumFlags++;
	}

	// Set up to parse for real
	CBinaryParser Parser2(cbBin,lpBin);
	if (efExtendedFlags.ulNumFlags && efExtendedFlags.ulNumFlags < ULONG_MAX/sizeof(ExtendedFlagStruct))
		efExtendedFlags.pefExtendedFlags = new ExtendedFlagStruct[efExtendedFlags.ulNumFlags];

	if (efExtendedFlags.pefExtendedFlags)
	{
		memset(efExtendedFlags.pefExtendedFlags,0,sizeof(ExtendedFlagStruct)*efExtendedFlags.ulNumFlags);
		ULONG i = 0;
		BOOL bBadData = false;

		for (i = 0 ; i < efExtendedFlags.ulNumFlags ; i++)
		{
			Parser2.GetBYTE(&efExtendedFlags.pefExtendedFlags[i].Id);
			Parser2.GetBYTE(&efExtendedFlags.pefExtendedFlags[i].Cb);

			// If the structure says there's more bytes than remaining buffer, we're done parsing.
			if (Parser2.RemainingBytes() < efExtendedFlags.pefExtendedFlags[i].Cb)
			{
				efExtendedFlags.ulNumFlags = i;
				break;
			}

			switch (efExtendedFlags.pefExtendedFlags[i].Id)
			{
			case EFPB_FLAGS:
				if (efExtendedFlags.pefExtendedFlags[i].Cb == sizeof(DWORD))
					Parser2.GetDWORD(&efExtendedFlags.pefExtendedFlags[i].Data.ExtendedFlags);
				else
					bBadData = true;
				break;
			case EFPB_CLSIDID:
				if (efExtendedFlags.pefExtendedFlags[i].Cb == sizeof(GUID))
					Parser2.GetBYTESNoAlloc(sizeof(GUID),(LPBYTE)&efExtendedFlags.pefExtendedFlags[i].Data.SearchFolderID);
				else
					bBadData = true;
				break;
			case EFPB_SFTAG:
				if (efExtendedFlags.pefExtendedFlags[i].Cb == sizeof(DWORD))
					Parser2.GetDWORD(&efExtendedFlags.pefExtendedFlags[i].Data.SearchFolderTag);
				else
					bBadData = true;
				break;
			case EFPB_TODO_VERSION:
				if (efExtendedFlags.pefExtendedFlags[i].Cb == sizeof(DWORD))
					Parser2.GetDWORD(&efExtendedFlags.pefExtendedFlags[i].Data.ToDoFolderVersion);
				else
					bBadData = true;
				break;
			default:
				Parser2.GetBYTES(efExtendedFlags.pefExtendedFlags[i].Cb,&efExtendedFlags.pefExtendedFlags[i].lpUnknownData);
				break;
			}

			// If we encountered a bad flag, stop parsing
			if (bBadData)
			{
				efExtendedFlags.ulNumFlags = i;
				break;
			}
		}
	}

	// Junk data remains
	if (Parser2.RemainingBytes() > 0)
	{
		efExtendedFlags.JunkDataSize = Parser2.RemainingBytes();
		Parser2.GetBYTES(efExtendedFlags.JunkDataSize,&efExtendedFlags.JunkData);
	}

	ExtendedFlagsStruct* pefExtendedFlags = new ExtendedFlagsStruct;
	if (pefExtendedFlags)
	{
		*pefExtendedFlags = efExtendedFlags;
	}

	return pefExtendedFlags;
}

void DeleteExtendedFlagsStruct(ExtendedFlagsStruct* pefExtendedFlags)
{
	if (!pefExtendedFlags) return;
	ULONG i = 0;
	pefExtendedFlags->ulNumFlags;
	for (i = 0 ; i < pefExtendedFlags->ulNumFlags ; i++)
	{
		delete[] pefExtendedFlags->pefExtendedFlags[i].lpUnknownData;
	}
	delete[] pefExtendedFlags->pefExtendedFlags;
	delete[] pefExtendedFlags->JunkData;
	delete pefExtendedFlags;
}

// result allocated with new, clean up with delete[]
LPTSTR ExtendedFlagsStructToString(ExtendedFlagsStruct* pefExtendedFlags)
{
	if (!pefExtendedFlags) return NULL;

	CString szExtendedFlags;
	CString szTmp;
	SBinary sBin = {0};

	szExtendedFlags.FormatMessage(IDS_EXTENDEDFLAGSHEADER,pefExtendedFlags->ulNumFlags);

	ULONG i = 0;
	for (i = 0 ; i < pefExtendedFlags->ulNumFlags ; i++)
	{
		LPTSTR szFlags = NULL;
		InterpretFlags(flagExtendedFolderFlagType,pefExtendedFlags->pefExtendedFlags[i].Id,&szFlags);
		szTmp.FormatMessage(IDS_EXTENDEDFLAGID,
			pefExtendedFlags->pefExtendedFlags[i].Id,szFlags,
			pefExtendedFlags->pefExtendedFlags[i].Cb);
		delete[] szFlags;
		szFlags = NULL;
		szExtendedFlags += szTmp;

		switch (pefExtendedFlags->pefExtendedFlags[i].Id)
		{
		case EFPB_FLAGS:
			InterpretFlags(flagExtendedFolderFlag,pefExtendedFlags->pefExtendedFlags[i].Data.ExtendedFlags,&szFlags);
			szTmp.FormatMessage(IDS_EXTENDEDFLAGDATAFLAG,pefExtendedFlags->pefExtendedFlags[i].Data.ExtendedFlags,szFlags);
			delete[] szFlags;
			szFlags = NULL;
			szExtendedFlags += szTmp;
			break;
		case EFPB_CLSIDID:
			szFlags = GUIDToString(&pefExtendedFlags->pefExtendedFlags[i].Data.SearchFolderID);
			szTmp.FormatMessage(IDS_EXTENDEDFLAGDATASFID,szFlags);
			delete[] szFlags;
			szFlags = NULL;
			szExtendedFlags += szTmp;
			break;
		case EFPB_SFTAG:
			szTmp.FormatMessage(IDS_EXTENDEDFLAGDATASFTAG,
				pefExtendedFlags->pefExtendedFlags[i].Data.SearchFolderTag);
			szExtendedFlags += szTmp;
			break;
		case EFPB_TODO_VERSION:
			szTmp.FormatMessage(IDS_EXTENDEDFLAGDATATODOVERSION,pefExtendedFlags->pefExtendedFlags[i].Data.ToDoFolderVersion);
			szExtendedFlags += szTmp;
			break;
		}
		if (pefExtendedFlags->pefExtendedFlags[i].lpUnknownData)
		{
			szTmp.FormatMessage(IDS_EXTENDEDFLAGUNKNOWN);
			szExtendedFlags += szTmp;
			sBin.cb = pefExtendedFlags->pefExtendedFlags[i].Cb;
			sBin.lpb = pefExtendedFlags->pefExtendedFlags[i].lpUnknownData;
			szExtendedFlags += BinToHexString(&sBin,true);
		}
	}

	szExtendedFlags += JunkDataToString(pefExtendedFlags->JunkDataSize,pefExtendedFlags->JunkData);

	return CStringToString(szExtendedFlags);
}

void TimeZoneToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	TimeZoneStruct* ptzTimeZone = BinToTimeZoneStruct(myBin.cb,myBin.lpb);
	if (ptzTimeZone)
	{
		*lpszResultString = TimeZoneStructToString(ptzTimeZone);
		DeleteTimeZoneStruct(ptzTimeZone);
	}
}

// Allocates return value with new. Clean up with DeleteTimeZoneStruct.
TimeZoneStruct* BinToTimeZoneStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	TimeZoneStruct tzTimeZone = {0};
	CBinaryParser Parser(cbBin,lpBin);

	Parser.GetDWORD(&tzTimeZone.lBias);
	Parser.GetDWORD(&tzTimeZone.lStandardBias);
	Parser.GetDWORD(&tzTimeZone.lDaylightBias);
	Parser.GetWORD(&tzTimeZone.wStandardYear);
	Parser.GetWORD(&tzTimeZone.stStandardDate.wYear);
	Parser.GetWORD(&tzTimeZone.stStandardDate.wMonth);
	Parser.GetWORD(&tzTimeZone.stStandardDate.wDayOfWeek);
	Parser.GetWORD(&tzTimeZone.stStandardDate.wDay);
	Parser.GetWORD(&tzTimeZone.stStandardDate.wHour);
	Parser.GetWORD(&tzTimeZone.stStandardDate.wMinute);
	Parser.GetWORD(&tzTimeZone.stStandardDate.wSecond);
	Parser.GetWORD(&tzTimeZone.stStandardDate.wMilliseconds);
	Parser.GetWORD(&tzTimeZone.wDaylightDate);
	Parser.GetWORD(&tzTimeZone.stDaylightDate.wYear);
	Parser.GetWORD(&tzTimeZone.stDaylightDate.wMonth);
	Parser.GetWORD(&tzTimeZone.stDaylightDate.wDayOfWeek);
	Parser.GetWORD(&tzTimeZone.stDaylightDate.wDay);
	Parser.GetWORD(&tzTimeZone.stDaylightDate.wHour);
	Parser.GetWORD(&tzTimeZone.stDaylightDate.wMinute);
	Parser.GetWORD(&tzTimeZone.stDaylightDate.wSecond);
	Parser.GetWORD(&tzTimeZone.stDaylightDate.wMilliseconds);

	// Junk data remains
	if (Parser.RemainingBytes() > 0)
	{
		tzTimeZone.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(tzTimeZone.JunkDataSize,&tzTimeZone.JunkData);
	}

	TimeZoneStruct* ptzTimeZone = new TimeZoneStruct;
	if (ptzTimeZone)
	{
		*ptzTimeZone = tzTimeZone;
	}

	return ptzTimeZone;
}

void DeleteTimeZoneStruct(TimeZoneStruct* ptzTimeZone)
{
	if (!ptzTimeZone) return;
	delete ptzTimeZone;
}

// result allocated with new, clean up with delete[]
LPTSTR TimeZoneStructToString(TimeZoneStruct* ptzTimeZone)
{
	if (!ptzTimeZone) return NULL;

	CString szTimeZone;

	szTimeZone.FormatMessage(IDS_TIMEZONE,
		ptzTimeZone->lBias,
		ptzTimeZone->lStandardBias,
		ptzTimeZone->lDaylightBias,
		ptzTimeZone->wStandardYear,
		ptzTimeZone->stStandardDate.wYear,
		ptzTimeZone->stStandardDate.wMonth,
		ptzTimeZone->stStandardDate.wDayOfWeek,
		ptzTimeZone->stStandardDate.wDay,
		ptzTimeZone->stStandardDate.wHour,
		ptzTimeZone->stStandardDate.wMinute,
		ptzTimeZone->stStandardDate.wSecond,
		ptzTimeZone->stStandardDate.wMilliseconds,
		ptzTimeZone->wDaylightDate,
		ptzTimeZone->stDaylightDate.wYear,
		ptzTimeZone->stDaylightDate.wMonth,
		ptzTimeZone->stDaylightDate.wDayOfWeek,
		ptzTimeZone->stDaylightDate.wDay,
		ptzTimeZone->stDaylightDate.wHour,
		ptzTimeZone->stDaylightDate.wMinute,
		ptzTimeZone->stDaylightDate.wSecond,
		ptzTimeZone->stDaylightDate.wMilliseconds);

	szTimeZone += JunkDataToString(ptzTimeZone->JunkDataSize,ptzTimeZone->JunkData);

	return CStringToString(szTimeZone);
}


void TimeZoneDefinitionToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	TimeZoneDefinitionStruct* ptzdTimeZoneDefinition = BinToTimeZoneDefinitionStruct(myBin.cb,myBin.lpb);
	if (ptzdTimeZoneDefinition)
	{
		*lpszResultString = TimeZoneDefinitionStructToString(ptzdTimeZoneDefinition);
		DeleteTimeZoneDefinitionStruct(ptzdTimeZoneDefinition);
	}
}

// Allocates return value with new. Clean up with DeleteTimeZoneDefinitionStruct.
TimeZoneDefinitionStruct* BinToTimeZoneDefinitionStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	TimeZoneDefinitionStruct tzdTimeZoneDefinition = {0};
	CBinaryParser Parser(cbBin,lpBin);

	Parser.GetBYTE(&tzdTimeZoneDefinition.bMajorVersion);
	Parser.GetBYTE(&tzdTimeZoneDefinition.bMinorVersion);
	Parser.GetWORD(&tzdTimeZoneDefinition.cbHeader);
	Parser.GetWORD(&tzdTimeZoneDefinition.wReserved);
	Parser.GetWORD(&tzdTimeZoneDefinition.cchKeyName);
	Parser.GetStringW(tzdTimeZoneDefinition.cchKeyName,&tzdTimeZoneDefinition.szKeyName);
	Parser.GetWORD(&tzdTimeZoneDefinition.cRules);

	if (tzdTimeZoneDefinition.cRules && tzdTimeZoneDefinition.cRules < TZ_MAX_RULES)
		tzdTimeZoneDefinition.lpTZRule = new TZRule[tzdTimeZoneDefinition.cRules];

	if (tzdTimeZoneDefinition.lpTZRule)
	{
		memset(tzdTimeZoneDefinition.lpTZRule,0,sizeof(TZRule) * tzdTimeZoneDefinition.cRules);
		ULONG i = 0;
		for (i = 0; i < tzdTimeZoneDefinition.cRules ; i++)
		{
			Parser.GetBYTE(&tzdTimeZoneDefinition.lpTZRule[i].bMajorVersion);
			Parser.GetBYTE(&tzdTimeZoneDefinition.lpTZRule[i].bMinorVersion);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].wReserved);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].wTZRuleFlags);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].wYear);
			Parser.GetBYTESNoAlloc(14,tzdTimeZoneDefinition.lpTZRule[i].X);
			Parser.GetDWORD(&tzdTimeZoneDefinition.lpTZRule[i].lBias);
			Parser.GetDWORD(&tzdTimeZoneDefinition.lpTZRule[i].lStandardBias);
			Parser.GetDWORD(&tzdTimeZoneDefinition.lpTZRule[i].lDaylightBias);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stStandardDate.wYear);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stStandardDate.wMonth);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stStandardDate.wDayOfWeek);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stStandardDate.wDay);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stStandardDate.wHour);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stStandardDate.wMinute);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stStandardDate.wSecond);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stStandardDate.wMilliseconds);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stDaylightDate.wYear);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stDaylightDate.wMonth);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stDaylightDate.wDayOfWeek);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stDaylightDate.wDay);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stDaylightDate.wHour);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stDaylightDate.wMinute);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stDaylightDate.wSecond);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].stDaylightDate.wMilliseconds);
		}
	}

	// Junk data remains
	if (Parser.RemainingBytes() > 0)
	{
		tzdTimeZoneDefinition.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(tzdTimeZoneDefinition.JunkDataSize,&tzdTimeZoneDefinition.JunkData);
	}

	TimeZoneDefinitionStruct* ptzdTimeZoneDefinition = new TimeZoneDefinitionStruct;
	if (ptzdTimeZoneDefinition)
	{
		*ptzdTimeZoneDefinition = tzdTimeZoneDefinition;
	}

	return ptzdTimeZoneDefinition;
}

void DeleteTimeZoneDefinitionStruct(TimeZoneDefinitionStruct* ptzdTimeZoneDefinition)
{
	if (!ptzdTimeZoneDefinition) return;
	delete[] ptzdTimeZoneDefinition->szKeyName;
	delete[] ptzdTimeZoneDefinition->lpTZRule;
	delete ptzdTimeZoneDefinition;
}

// result allocated with new, clean up with delete[]
LPTSTR TimeZoneDefinitionStructToString(TimeZoneDefinitionStruct* ptzdTimeZoneDefinition)
{
	if (!ptzdTimeZoneDefinition) return NULL;

	CString szTimeZoneDefinition;
	CString szTmp;
	HRESULT hRes = S_OK;

	szTimeZoneDefinition.FormatMessage(IDS_TIMEZONEDEFINITION,
		ptzdTimeZoneDefinition->bMajorVersion,
		ptzdTimeZoneDefinition->bMinorVersion,
		ptzdTimeZoneDefinition->cbHeader,
		ptzdTimeZoneDefinition->wReserved,
		ptzdTimeZoneDefinition->cchKeyName,
		ptzdTimeZoneDefinition->szKeyName,
		ptzdTimeZoneDefinition->cRules);

	if (ptzdTimeZoneDefinition->cRules && ptzdTimeZoneDefinition->lpTZRule)
	{
		CString szTZRule;

		WORD i = 0;

		for (i = 0;i<ptzdTimeZoneDefinition->cRules;i++)
		{
			LPTSTR szFlags = NULL;
			EC_H(InterpretFlags(flagTZRule, ptzdTimeZoneDefinition->lpTZRule[i].wTZRuleFlags, &szFlags));
			szTZRule.FormatMessage(IDS_TZRULEHEADER,
				i,
				ptzdTimeZoneDefinition->lpTZRule[i].bMajorVersion,
				ptzdTimeZoneDefinition->lpTZRule[i].bMinorVersion,
				ptzdTimeZoneDefinition->lpTZRule[i].wReserved,
				ptzdTimeZoneDefinition->lpTZRule[i].wTZRuleFlags,
				szFlags,
				ptzdTimeZoneDefinition->lpTZRule[i].wYear);
			delete[] szFlags;
			szFlags = NULL;

			SBinary sBin = {0};
			sBin.cb = 14;
			sBin.lpb = ptzdTimeZoneDefinition->lpTZRule[i].X;
			szTZRule += BinToHexString(&sBin,true);

			szTmp.FormatMessage(IDS_TZRULEFOOTER,
				i,
				ptzdTimeZoneDefinition->lpTZRule[i].lBias,
				ptzdTimeZoneDefinition->lpTZRule[i].lStandardBias,
				ptzdTimeZoneDefinition->lpTZRule[i].lDaylightBias,
				ptzdTimeZoneDefinition->lpTZRule[i].stStandardDate.wYear,
				ptzdTimeZoneDefinition->lpTZRule[i].stStandardDate.wMonth,
				ptzdTimeZoneDefinition->lpTZRule[i].stStandardDate.wDayOfWeek,
				ptzdTimeZoneDefinition->lpTZRule[i].stStandardDate.wDay,
				ptzdTimeZoneDefinition->lpTZRule[i].stStandardDate.wHour,
				ptzdTimeZoneDefinition->lpTZRule[i].stStandardDate.wMinute,
				ptzdTimeZoneDefinition->lpTZRule[i].stStandardDate.wSecond,
				ptzdTimeZoneDefinition->lpTZRule[i].stStandardDate.wMilliseconds,
				ptzdTimeZoneDefinition->lpTZRule[i].stDaylightDate.wYear,
				ptzdTimeZoneDefinition->lpTZRule[i].stDaylightDate.wMonth,
				ptzdTimeZoneDefinition->lpTZRule[i].stDaylightDate.wDayOfWeek,
				ptzdTimeZoneDefinition->lpTZRule[i].stDaylightDate.wDay,
				ptzdTimeZoneDefinition->lpTZRule[i].stDaylightDate.wHour,
				ptzdTimeZoneDefinition->lpTZRule[i].stDaylightDate.wMinute,
				ptzdTimeZoneDefinition->lpTZRule[i].stDaylightDate.wSecond,
				ptzdTimeZoneDefinition->lpTZRule[i].stDaylightDate.wMilliseconds);
			szTZRule += szTmp;

			szTimeZoneDefinition += szTZRule;
		}
	}

	szTimeZoneDefinition += JunkDataToString(ptzdTimeZoneDefinition->JunkDataSize,ptzdTimeZoneDefinition->JunkData);

	return CStringToString(szTimeZoneDefinition);
}


void ReportTagToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	ReportTagStruct* prtReportTag = BinToReportTagStruct(myBin.cb,myBin.lpb);
	if (prtReportTag)
	{
		*lpszResultString = ReportTagStructToString(prtReportTag);
		DeleteReportTagStruct(prtReportTag);
	}
}

// Allocates return value with new. Clean up with DeleteReportTagStruct.
ReportTagStruct* BinToReportTagStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	ReportTagStruct rtReportTag = {0};
	CBinaryParser Parser(cbBin,lpBin);

	Parser.GetBYTESNoAlloc(sizeof(rtReportTag.Cookie),(LPBYTE) rtReportTag.Cookie);

	// Version is big endian, so we have to read individual bytes
	WORD hiWord = NULL;
	WORD loWord = NULL;
	Parser.GetWORD(&hiWord);
	Parser.GetWORD(&loWord);
	rtReportTag.Version = (hiWord << 16) | loWord;

	Parser.GetDWORD(&rtReportTag.cbStoreEntryID);
	if (rtReportTag.cbStoreEntryID)
	{
		Parser.GetBYTES(rtReportTag.cbStoreEntryID,&rtReportTag.lpStoreEntryID);
	}
	Parser.GetDWORD(&rtReportTag.cbFolderEntryID);
	if (rtReportTag.cbFolderEntryID)
	{
		Parser.GetBYTES(rtReportTag.cbFolderEntryID,&rtReportTag.lpFolderEntryID);
	}
	Parser.GetDWORD(&rtReportTag.cbMessageEntryID);
	if (rtReportTag.cbMessageEntryID)
	{
		Parser.GetBYTES(rtReportTag.cbMessageEntryID,&rtReportTag.lpMessageEntryID);
	}
	Parser.GetDWORD(&rtReportTag.cbSearchFolderEntryID);
	if (rtReportTag.cbSearchFolderEntryID)
	{
		Parser.GetBYTES(rtReportTag.cbSearchFolderEntryID,&rtReportTag.lpSearchFolderEntryID);
	}
	Parser.GetDWORD(&rtReportTag.cbMessageSearchKey);
	if (rtReportTag.cbMessageSearchKey)
	{
		Parser.GetBYTES(rtReportTag.cbMessageSearchKey,&rtReportTag.lpMessageSearchKey);
	}
	Parser.GetDWORD(&rtReportTag.cchAnsiText);
	if (rtReportTag.cchAnsiText)
	{
		Parser.GetStringA(rtReportTag.cchAnsiText,&rtReportTag.lpszAnsiText);
	}

	// Junk data remains
	if (Parser.RemainingBytes() > 0)
	{
		rtReportTag.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(rtReportTag.JunkDataSize,&rtReportTag.JunkData);
	}

	ReportTagStruct* prtReportTag = new ReportTagStruct;
	if (prtReportTag)
	{
		*prtReportTag = rtReportTag;
	}

	return prtReportTag;
}

void DeleteReportTagStruct(ReportTagStruct* prtReportTag)
{
	if (!prtReportTag) return;
	delete[] prtReportTag->lpStoreEntryID;
	delete[] prtReportTag->lpFolderEntryID;
	delete[] prtReportTag->lpMessageEntryID;
	delete[] prtReportTag->lpSearchFolderEntryID;
	delete[] prtReportTag->lpMessageSearchKey;
	delete[] prtReportTag->lpszAnsiText;
	delete[] prtReportTag->JunkData;
	delete prtReportTag;
}

// result allocated with new, clean up with delete[]
LPTSTR ReportTagStructToString(ReportTagStruct* prtReportTag)
{
	if (!prtReportTag) return NULL;

	CString szReportTag;
	CString szTmp;
	HRESULT hRes = S_OK;

	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagReportTagVersion, prtReportTag->Version, &szFlags));
	szTmp.FormatMessage(IDS_REPORTTAGHEADER,
		prtReportTag->Cookie,
		prtReportTag->Version,
		szFlags);
	delete[] szFlags;
	szFlags = NULL;
	szReportTag += szTmp;

	if (prtReportTag->cbStoreEntryID)
	{
		szTmp.FormatMessage(IDS_REPORTTAGSTOREEID);
		szReportTag += szTmp;
		SBinary sBin = {0};
		sBin.cb = prtReportTag->cbStoreEntryID;
		sBin.lpb = prtReportTag->lpStoreEntryID;
		szReportTag += BinToHexString(&sBin,true);
	}

	if (prtReportTag->cbFolderEntryID)
	{
		szTmp.FormatMessage(IDS_REPORTTAGFOLDEREID);
		szReportTag += szTmp;
		SBinary sBin = {0};
		sBin.cb = prtReportTag->cbFolderEntryID;
		sBin.lpb = prtReportTag->lpFolderEntryID;
		szReportTag += BinToHexString(&sBin,true);
	}

	if (prtReportTag->cbMessageEntryID)
	{
		szTmp.FormatMessage(IDS_REPORTTAGMESSAGEEID);
		szReportTag += szTmp;
		SBinary sBin = {0};
		sBin.cb = prtReportTag->cbMessageEntryID;
		sBin.lpb = prtReportTag->lpMessageEntryID;
		szReportTag += BinToHexString(&sBin,true);
	}

	if (prtReportTag->cbSearchFolderEntryID)
	{
		szTmp.FormatMessage(IDS_REPORTTAGSFEID);
		szReportTag += szTmp;
		SBinary sBin = {0};
		sBin.cb = prtReportTag->cbSearchFolderEntryID;
		sBin.lpb = prtReportTag->lpSearchFolderEntryID;
		szReportTag += BinToHexString(&sBin,true);
	}

	if (prtReportTag->cbMessageSearchKey)
	{
		szTmp.FormatMessage(IDS_REPORTTAGMESSAGEKEY);
		szReportTag += szTmp;
		SBinary sBin = {0};
		sBin.cb = prtReportTag->cbMessageSearchKey;
		sBin.lpb = prtReportTag->lpMessageSearchKey;
		szReportTag += BinToHexString(&sBin,true);
	}

	if (prtReportTag->cchAnsiText)
	{
		szTmp.FormatMessage(IDS_REPORTTAGANSITEXT,
			prtReportTag->cchAnsiText,
			prtReportTag->lpszAnsiText?prtReportTag->lpszAnsiText:""); // STRING_OK
		szReportTag += szTmp;
	}

	szReportTag += JunkDataToString(prtReportTag->JunkDataSize,prtReportTag->JunkData);

	return CStringToString(szReportTag);
}

void ConversationIndexToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	ConversationIndexStruct* pciConversationIndex = BinToConversationIndexStruct(myBin.cb,myBin.lpb);
	if (pciConversationIndex)
	{
		*lpszResultString = ConversationIndexStructToString(pciConversationIndex);
		DeleteConversationIndexStruct(pciConversationIndex);
	}
}

// Allocates return value with new. Clean up with DeleteConversationIndexStruct.
ConversationIndexStruct* BinToConversationIndexStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	ConversationIndexStruct ciConversationIndex = {0};
	CBinaryParser Parser(cbBin,lpBin);

	Parser.GetBYTE(&ciConversationIndex.UnnamedByte);
	BYTE b1 = NULL;
	BYTE b2 = NULL;
	BYTE b3 = NULL;
	BYTE b4 = NULL;
	Parser.GetBYTE(&b1);
	Parser.GetBYTE(&b2);
	Parser.GetBYTE(&b3);
	ciConversationIndex.ftCurrent.dwHighDateTime = (b1 << 16) | (b2 << 8) | b3;
	Parser.GetBYTE(&b1);
	Parser.GetBYTE(&b2);
	ciConversationIndex.ftCurrent.dwLowDateTime = (b1 << 24) | (b2 << 16);
	Parser.GetBYTE(&b1);
	Parser.GetBYTE(&b2);
	Parser.GetBYTE(&b3);
	Parser.GetBYTE(&b4);
	ciConversationIndex.guid.Data1 = (b1 << 24) | (b2 << 16) | (b3 << 8) | b4;
	Parser.GetBYTE(&b1);
	Parser.GetBYTE(&b2);
	ciConversationIndex.guid.Data2 = (b1 << 8) | b2;
	Parser.GetBYTE(&b1);
	Parser.GetBYTE(&b2);
	ciConversationIndex.guid.Data3 = (b1 << 8) | b2;
	Parser.GetBYTESNoAlloc(sizeof(ciConversationIndex.guid.Data4),ciConversationIndex.guid.Data4);

	if (Parser.RemainingBytes() > 0)
	{
		ciConversationIndex.ulResponseLevels = (ULONG) Parser.RemainingBytes()/5; // Response levels consume 5 bytes each
	}

	if (ciConversationIndex.ulResponseLevels && ciConversationIndex.ulResponseLevels < ULONG_MAX/sizeof(ResponseLevelStruct))
		ciConversationIndex.lpResponseLevels = new ResponseLevelStruct[ciConversationIndex.ulResponseLevels];

	if (ciConversationIndex.lpResponseLevels)
	{
		memset(ciConversationIndex.lpResponseLevels,0,sizeof(ResponseLevelStruct) * ciConversationIndex.ulResponseLevels);
		ULONG i = 0;
		for (i = 0; i < ciConversationIndex.ulResponseLevels ; i++)
		{
			Parser.GetBYTE(&b1);
			Parser.GetBYTE(&b2);
			Parser.GetBYTE(&b3);
			Parser.GetBYTE(&b4);
			ciConversationIndex.lpResponseLevels[i].TimeDelta = (b1 << 24) | (b2 << 16) | (b3 << 8) | b4;
			if (ciConversationIndex.lpResponseLevels[i].TimeDelta & 0x80000000)
			{
				ciConversationIndex.lpResponseLevels[i].TimeDelta = ciConversationIndex.lpResponseLevels[i].TimeDelta & ~0x80000000;
				ciConversationIndex.lpResponseLevels[i].DeltaCode = true;
			}
			Parser.GetBYTE(&b1);
			ciConversationIndex.lpResponseLevels[i].Random = b1 >> 4;
			ciConversationIndex.lpResponseLevels[i].ResponseLevel = b1 & 0xf;
		}
	}

	// Junk data remains
	if (Parser.RemainingBytes() > 0)
	{
		ciConversationIndex.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(ciConversationIndex.JunkDataSize,&ciConversationIndex.JunkData);
	}

	ConversationIndexStruct* pciConversationIndex = new ConversationIndexStruct;
	if (pciConversationIndex)
	{
		*pciConversationIndex = ciConversationIndex;
	}

	return pciConversationIndex;
}

void DeleteConversationIndexStruct(ConversationIndexStruct* pciConversationIndex)
{
	if (!pciConversationIndex) return;
	delete[] pciConversationIndex->lpResponseLevels;
	delete[] pciConversationIndex->JunkData;
	delete pciConversationIndex;
}

// result allocated with new, clean up with delete[]
LPTSTR ConversationIndexStructToString(ConversationIndexStruct* pciConversationIndex)
{
	if (!pciConversationIndex) return NULL;

	CString szConversationIndex;
	CString szTmp;

	CString PropString;
	FileTimeToString(&pciConversationIndex->ftCurrent,&PropString,NULL);
	LPTSTR szGUID = GUIDToString(&pciConversationIndex->guid);
	szTmp.FormatMessage(IDS_CONVERSATIONINDEXHEADER,
		pciConversationIndex->UnnamedByte,
		pciConversationIndex->ftCurrent.dwLowDateTime,
		pciConversationIndex->ftCurrent.dwHighDateTime,
		PropString,
		szGUID);
	szConversationIndex += szTmp;
	delete[] szGUID;

	if (pciConversationIndex->ulResponseLevels)
	{
		ULONG i = 0;
		for (i = 0 ; i < pciConversationIndex->ulResponseLevels ; i++)
		{
			szTmp.FormatMessage(IDS_CONVERSATIONINDEXRESPONSELEVEL,
				i,pciConversationIndex->lpResponseLevels[i].DeltaCode,
				pciConversationIndex->lpResponseLevels[i].TimeDelta,
				pciConversationIndex->lpResponseLevels[i].Random,
				pciConversationIndex->lpResponseLevels[i].ResponseLevel);
			szConversationIndex += szTmp;
		}
	}

	szConversationIndex += JunkDataToString(pciConversationIndex->JunkDataSize,pciConversationIndex->JunkData);

	return CStringToString(szConversationIndex);
}

void TaskAssignersToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	TaskAssignersStruct* ptaTaskAssigners = BinToTaskAssignersStruct(myBin.cb,myBin.lpb);
	if (ptaTaskAssigners)
	{
		*lpszResultString = TaskAssignersStructToString(ptaTaskAssigners);
		DeleteTaskAssignersStruct(ptaTaskAssigners);
	}
}

// Allocates return value with new. Clean up with DeleteTaskAssignersStruct.
TaskAssignersStruct* BinToTaskAssignersStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	TaskAssignersStruct taTaskAssigners = {0};
	CBinaryParser Parser(cbBin,lpBin);

	Parser.GetDWORD(&taTaskAssigners.cAssigners);

	if (taTaskAssigners.cAssigners && taTaskAssigners.cAssigners < ULONG_MAX/sizeof(TaskAssignerStruct))
		taTaskAssigners.lpTaskAssigners = new TaskAssignerStruct[taTaskAssigners.cAssigners];

	if (taTaskAssigners.lpTaskAssigners)
	{
		memset(taTaskAssigners.lpTaskAssigners,0,sizeof(TaskAssignerStruct) * taTaskAssigners.cAssigners);
		DWORD i = 0;
		for (i = 0 ; i < taTaskAssigners.cAssigners ; i++)
		{
			Parser.GetDWORD(&taTaskAssigners.lpTaskAssigners[i].cbAssigner);
			LPBYTE lpbAssigner = NULL;
			Parser.GetBYTES(taTaskAssigners.lpTaskAssigners[i].cbAssigner,&lpbAssigner);
			if (lpbAssigner)
			{
				CBinaryParser AssignerParser(taTaskAssigners.lpTaskAssigners[i].cbAssigner,lpbAssigner);
				AssignerParser.GetDWORD(&taTaskAssigners.lpTaskAssigners[i].cbEntryID);
				AssignerParser.GetBYTES(taTaskAssigners.lpTaskAssigners[i].cbEntryID,&taTaskAssigners.lpTaskAssigners[i].lpEntryID);
				AssignerParser.GetStringA(&taTaskAssigners.lpTaskAssigners[i].szDisplayName);
				AssignerParser.GetStringW(&taTaskAssigners.lpTaskAssigners[i].wzDisplayName);

				// Junk data remains
				if (AssignerParser.RemainingBytes() > 0)
				{
					taTaskAssigners.lpTaskAssigners[i].JunkDataSize = AssignerParser.RemainingBytes();
					AssignerParser.GetBYTES(taTaskAssigners.lpTaskAssigners[i].JunkDataSize,&taTaskAssigners.lpTaskAssigners[i].JunkData);
				}
				delete[] lpbAssigner;
			}
		}
	}

	// Junk data remains
	if (Parser.RemainingBytes() > 0)
	{
		taTaskAssigners.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(taTaskAssigners.JunkDataSize,&taTaskAssigners.JunkData);
	}

	TaskAssignersStruct* ptaTaskAssigners = new TaskAssignersStruct;
	if (ptaTaskAssigners)
	{
		*ptaTaskAssigners = taTaskAssigners;
	}

	return ptaTaskAssigners;
}
void DeleteTaskAssignersStruct(TaskAssignersStruct* ptaTaskAssigners)
{
	if (!ptaTaskAssigners) return;
	DWORD i = 0;
	for (i = 0 ; i < ptaTaskAssigners->cAssigners ; i++)
	{
		delete[] ptaTaskAssigners->lpTaskAssigners[i].lpEntryID;
		delete[] ptaTaskAssigners->lpTaskAssigners[i].szDisplayName;
		delete[] ptaTaskAssigners->lpTaskAssigners[i].wzDisplayName;
		delete[] ptaTaskAssigners->lpTaskAssigners[i].JunkData;
	}
	delete[] ptaTaskAssigners->lpTaskAssigners;
	delete[] ptaTaskAssigners->JunkData;
	delete ptaTaskAssigners;
}
// result allocated with new, clean up with delete[]
LPTSTR TaskAssignersStructToString(TaskAssignersStruct* ptaTaskAssigners)
{
	if (!ptaTaskAssigners) return NULL;

	CString szTaskAssigners;
	CString szTmp;

	szTaskAssigners.FormatMessage(IDS_TASKASSIGNERSHEADER,
		ptaTaskAssigners->cAssigners);

	if (ptaTaskAssigners->cAssigners && ptaTaskAssigners->lpTaskAssigners)
	{
		DWORD i = 0;
		for (i = 0 ; i < ptaTaskAssigners->cAssigners ; i++)
		{
			szTmp.FormatMessage(IDS_TASKASSIGNEREID,
				i,
				ptaTaskAssigners->lpTaskAssigners[i].cbEntryID);
			szTaskAssigners += szTmp;
			if (ptaTaskAssigners->lpTaskAssigners[i].lpEntryID)
			{
				SBinary sBin = {0};
				sBin.cb = ptaTaskAssigners->lpTaskAssigners[i].cbEntryID;
				sBin.lpb = ptaTaskAssigners->lpTaskAssigners[i].lpEntryID;
				szTaskAssigners += BinToHexString(&sBin,true);
			}
			szTmp.FormatMessage(IDS_TASKASSIGNERNAME,
				ptaTaskAssigners->lpTaskAssigners[i].szDisplayName,
				ptaTaskAssigners->lpTaskAssigners[i].wzDisplayName);
			szTaskAssigners += szTmp;

			if (ptaTaskAssigners->lpTaskAssigners[i].JunkDataSize)
			{
				szTmp.FormatMessage(IDS_TASKASSIGNERJUNKDATA,
					ptaTaskAssigners->lpTaskAssigners[i].JunkDataSize);
				szTaskAssigners += szTmp;
				SBinary sBin = {0};
				sBin.cb = (ULONG) ptaTaskAssigners->lpTaskAssigners[i].JunkDataSize;
				sBin.lpb = ptaTaskAssigners->lpTaskAssigners[i].JunkData;
				szTaskAssigners += BinToHexString(&sBin,true);
			}
		}
	}

	szTaskAssigners += JunkDataToString(ptaTaskAssigners->JunkDataSize,ptaTaskAssigners->JunkData);

	return CStringToString(szTaskAssigners);
}

void GlobalObjectIdToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	GlobalObjectIdStruct* ptaGlobalObjectId = BinToGlobalObjectIdStruct(myBin.cb,myBin.lpb);
	if (ptaGlobalObjectId)
	{
		*lpszResultString = GlobalObjectIdStructToString(ptaGlobalObjectId);
		DeleteGlobalObjectIdStruct(ptaGlobalObjectId);
	}
}

// Allocates return value with new. Clean up with DeleteGlobalObjectIdStruct.
GlobalObjectIdStruct* BinToGlobalObjectIdStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	GlobalObjectIdStruct goidGlobalObjectId = {0};
	CBinaryParser Parser(cbBin,lpBin);

	Parser.GetBYTESNoAlloc(sizeof(goidGlobalObjectId.Id),(LPBYTE) &goidGlobalObjectId.Id);
	BYTE b1 = NULL;
	BYTE b2 = NULL;
	Parser.GetBYTE(&b1);
	Parser.GetBYTE(&b2);
	goidGlobalObjectId.Year = (b1 << 8) | b2;
	Parser.GetBYTE(&goidGlobalObjectId.Month);
	Parser.GetBYTE(&goidGlobalObjectId.Day);
	Parser.GetLARGE_INTEGER((LARGE_INTEGER*) &goidGlobalObjectId.CreationTime);
	Parser.GetLARGE_INTEGER(&goidGlobalObjectId.X);
	Parser.GetDWORD(&goidGlobalObjectId.dwSize);
	Parser.GetBYTES(goidGlobalObjectId.dwSize,&goidGlobalObjectId.lpData);

	// Junk data remains
	if (Parser.RemainingBytes() > 0)
	{
		goidGlobalObjectId.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(goidGlobalObjectId.JunkDataSize,&goidGlobalObjectId.JunkData);
	}

	GlobalObjectIdStruct* pgoidGlobalObjectId = new GlobalObjectIdStruct;
	if (pgoidGlobalObjectId)
	{
		*pgoidGlobalObjectId = goidGlobalObjectId;
	}

	return pgoidGlobalObjectId;
}

void DeleteGlobalObjectIdStruct(GlobalObjectIdStruct* pgoidGlobalObjectId)
{
	if (!pgoidGlobalObjectId) return;
	delete[] pgoidGlobalObjectId->lpData;
	delete[] pgoidGlobalObjectId->JunkData;
	delete pgoidGlobalObjectId;
}

// result allocated with new, clean up with delete[]
LPTSTR GlobalObjectIdStructToString(GlobalObjectIdStruct* pgoidGlobalObjectId)
{
	if (!pgoidGlobalObjectId) return NULL;

	CString szGlobalObjectId;
	CString szTmp;
	HRESULT hRes = S_OK;

	szGlobalObjectId.FormatMessage(IDS_GLOBALOBJECTIDHEADER);

	SBinary sBin = {0};
	sBin.cb = sizeof(pgoidGlobalObjectId->Id);
	sBin.lpb = pgoidGlobalObjectId->Id;
	szGlobalObjectId += BinToHexString(&sBin,true);

	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagGlobalObjectIdMonth, pgoidGlobalObjectId->Month, &szFlags));

	CString PropString;
	FileTimeToString(&pgoidGlobalObjectId->CreationTime,&PropString,NULL);
	szTmp.FormatMessage(IDS_GLOBALOBJECTIDDATA1,
		pgoidGlobalObjectId->Year,
		pgoidGlobalObjectId->Month,szFlags,
		pgoidGlobalObjectId->Day,
		pgoidGlobalObjectId->CreationTime.dwHighDateTime,pgoidGlobalObjectId->CreationTime.dwLowDateTime,PropString,
		pgoidGlobalObjectId->X.HighPart,pgoidGlobalObjectId->X.LowPart,
		pgoidGlobalObjectId->dwSize);
	delete[] szFlags;
	szFlags = NULL;
	szGlobalObjectId += szTmp;

	if (pgoidGlobalObjectId->dwSize && pgoidGlobalObjectId->lpData)
	{
		szTmp.FormatMessage(IDS_GLOBALOBJECTIDDATA2);
		szGlobalObjectId += szTmp;
		sBin.cb = pgoidGlobalObjectId->dwSize;
		sBin.lpb = pgoidGlobalObjectId->lpData;
		szGlobalObjectId += BinToHexString(&sBin,true);
	}

	szGlobalObjectId += JunkDataToString(pgoidGlobalObjectId->JunkDataSize,pgoidGlobalObjectId->JunkData);

	return CStringToString(szGlobalObjectId);
}

void OneOffEntryIdToString(SBinary myBin, LPTSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	OneOffEntryIdStruct* pooeidOneOffEntryId = BinToOneOffEntryIdStruct(myBin.cb,myBin.lpb);
	if (pooeidOneOffEntryId)
	{
		*lpszResultString = OneOffEntryIdStructToString(pooeidOneOffEntryId);
		DeleteOneOffEntryIdStruct(pooeidOneOffEntryId);
	}
}

// Allocates return value with new. Clean up with DeleteOneOffEntryIdStruct.
OneOffEntryIdStruct* BinToOneOffEntryIdStruct(ULONG cbBin, LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	OneOffEntryIdStruct ooeidOneOffEntryId = {0};
	CBinaryParser Parser(cbBin,lpBin);
	Parser.GetDWORD(&ooeidOneOffEntryId.dwFlags);
	Parser.GetBYTESNoAlloc(sizeof(ooeidOneOffEntryId.ProviderUID),(LPBYTE) &ooeidOneOffEntryId.ProviderUID);
	Parser.GetDWORD(&ooeidOneOffEntryId.dwBitmask);

	if (MAPI_UNICODE & ooeidOneOffEntryId.dwBitmask)
	{
		Parser.GetStringW(&ooeidOneOffEntryId.Strings.Unicode.szDisplayName);
		Parser.GetStringW(&ooeidOneOffEntryId.Strings.Unicode.szAddressType);
		Parser.GetStringW(&ooeidOneOffEntryId.Strings.Unicode.szEmailAddress);
	}
	else
	{
		Parser.GetStringA(&ooeidOneOffEntryId.Strings.ANSI.szDisplayName);
		Parser.GetStringA(&ooeidOneOffEntryId.Strings.ANSI.szAddressType);
		Parser.GetStringA(&ooeidOneOffEntryId.Strings.ANSI.szEmailAddress);
	}

	// Junk data remains
	if (Parser.RemainingBytes() > 0)
	{
		ooeidOneOffEntryId.JunkDataSize = Parser.RemainingBytes();
		Parser.GetBYTES(ooeidOneOffEntryId.JunkDataSize,&ooeidOneOffEntryId.JunkData);
	}

	OneOffEntryIdStruct* pooeidOneOffEntryId = new OneOffEntryIdStruct;
	if (pooeidOneOffEntryId)
	{
		*pooeidOneOffEntryId = ooeidOneOffEntryId;
	}

	return pooeidOneOffEntryId;
}

void DeleteOneOffEntryIdStruct(OneOffEntryIdStruct* pooeidOneOffEntryId)
{
	if (!pooeidOneOffEntryId) return;
	if (MAPI_UNICODE & pooeidOneOffEntryId->dwBitmask)
	{
		delete[] pooeidOneOffEntryId->Strings.Unicode.szDisplayName;
		delete[] pooeidOneOffEntryId->Strings.Unicode.szAddressType;
		delete[] pooeidOneOffEntryId->Strings.Unicode.szEmailAddress;
	}
	else
	{
		delete[] pooeidOneOffEntryId->Strings.ANSI.szDisplayName;
		delete[] pooeidOneOffEntryId->Strings.ANSI.szAddressType;
		delete[] pooeidOneOffEntryId->Strings.ANSI.szEmailAddress;
	}
	delete pooeidOneOffEntryId;
}

// result allocated with new, clean up with delete[]
LPTSTR OneOffEntryIdStructToString(OneOffEntryIdStruct* pooeidOneOffEntryId)
{
	if (!pooeidOneOffEntryId) return NULL;

	CString szOneOffEntryId;
	CString szTmp;
	HRESULT hRes = S_OK;

	szOneOffEntryId.FormatMessage(IDS_ONEOFFENTRYIDHEADER,
		pooeidOneOffEntryId->dwFlags);

	SBinary sBin = {0};
	sBin.cb = sizeof(pooeidOneOffEntryId->ProviderUID);
	sBin.lpb = pooeidOneOffEntryId->ProviderUID;
	szOneOffEntryId += BinToHexString(&sBin,true);

	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagOneOffEntryId, pooeidOneOffEntryId->dwBitmask, &szFlags));

	if (MAPI_UNICODE & pooeidOneOffEntryId->dwBitmask)
	{
		szTmp.FormatMessage(IDS_ONEOFFENTRYIDFOOTERUNICODE,
			pooeidOneOffEntryId->dwBitmask,szFlags,
			pooeidOneOffEntryId->Strings.Unicode.szDisplayName,
			pooeidOneOffEntryId->Strings.Unicode.szAddressType,
			pooeidOneOffEntryId->Strings.Unicode.szEmailAddress);
	}
	else
	{
		szTmp.FormatMessage(IDS_ONEOFFENTRYIDFOOTERANSI,
			pooeidOneOffEntryId->dwBitmask,szFlags,
			pooeidOneOffEntryId->Strings.ANSI.szDisplayName,
			pooeidOneOffEntryId->Strings.ANSI.szAddressType,
			pooeidOneOffEntryId->Strings.ANSI.szEmailAddress);
	}
	delete[] szFlags;
	szFlags = NULL;
	szOneOffEntryId += szTmp;

	szOneOffEntryId += JunkDataToString(pooeidOneOffEntryId->JunkDataSize,pooeidOneOffEntryId->JunkData);

	return CStringToString(szOneOffEntryId);
}
