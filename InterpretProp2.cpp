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

	InterpretFlags(ulPropID,lVal,szFlagString);
	return S_OK;
}

// Interprets a flag value according to a flag name and returns a string
// allocated with new
// Free the string with delete[]
// Will not return a string if the flag name is not recognized
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
		*szFlagString = new TCHAR[cchLen];

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

	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidTimeZoneStruct,PSETID_Appointment,stTZREG)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidApptTZDefStartDisplay,PSETID_Appointment,stTZDEFINITION)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidApptTZDefEndDisplay,PSETID_Appointment,stTZDEFINITION)
	NAMEDPROP_BINARY_STRUCTURE_ENTRY(dispidApptTZDefRecur,PSETID_Appointment,stTZDEFINITION)
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
				EC_H(InterpretFlags(lpProp,ulPropNameID,NULL,lpNameID?lpNameID->lpguid:NULL,lpszSmartView));
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
			lpStream->Stat(&StatInfo, STATFLAG_NONAME);//find out how much space we need

			// We're not going to try to support MASSIVE properties. 
			if (!StatInfo.cbSize.HighPart)
			{
				EC_H(MAPIAllocateBuffer(
					sizeof(SPropValue),
					(LPVOID*) lpPropArray));
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
	case stTZDEFINITION:
		{
			TZDEFINITION* ptzDefinition = BinToTZDEFINITION(myBin.cb,myBin.lpb);
			if (ptzDefinition)
			{
				szResultString = TZDEFINITIONToString(*ptzDefinition);
				delete[] ptzDefinition;
			}
		}
		break;
	case stTZREG:
		{
			TZREG* ptzReg = BinToTZREG(myBin.cb,myBin.lpb);
			if (ptzReg)
			{
				szResultString = TZREGToString(*ptzReg);
				delete ptzReg;
			}
		}
		break;
	case stSecurityDescriptor:
		{
			SDBinToString(myBin,lpMAPIProp,ulPropTag,&szResultString);
		}
		break;
	case stExtendedFolderFlags:
		{
			ExtendedFlagsBinToString(myBin,&szResultString);
		}
		break;
	}
	if (szResultString) *lpszResultString = szResultString;
}

void SDBinToString(SBinary myBin, LPMAPIPROP lpMAPIProp, ULONG ulPropTag, LPTSTR* lpszResultString)
{
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
	CString szResultString;
	LPTSTR szOut = NULL;

	LPBYTE lpCur = myBin.lpb;
	BOOL bFirstPass = true;

	for (;;)
	{
		// Check that we have room for an Id and Cb
		if (myBin.lpb + myBin.cb < lpCur + 2) break;
		if (!bFirstPass) szResultString += _T("\n"); // STRING_OK

		BYTE bID = *lpCur;
		lpCur += 1;

		CString szData;
		BOOL bDataHandled = false;

		LPTSTR szFlags = NULL;
		InterpretFlags(flagExtendedFolderFlagType,bID,&szFlags);
		szResultString += szFlags;
		szResultString += _T(": "); // STRING_OK
		delete[] szFlags;
		szFlags = NULL;

		ULONG cbProp = *lpCur;

		// Check that we have room for our data
		if (lpCur + 1 + cbProp > myBin.lpb + myBin.cb) 
		{
			// We don't. Break out of here and let the junk routine catch what remains
			break;
		}

		lpCur += 1;

		switch (bID)
		{
		case EFPB_FLAGS:
			{
				DWORD dwExFlags = NULL;
				if (cbProp == sizeof(DWORD)) 
				{
					dwExFlags = *((DWORD*)lpCur);
					InterpretFlags(flagExtendedFolderFlag,dwExFlags,&szFlags);
					szData.Format(_T("0x%08X = %s"),dwExFlags,szFlags); // STRING_OK
					delete[] szFlags;
					szFlags = NULL;
					bDataHandled = true;
				}
			}
			break;
		case EFPB_CLSIDID:
			{
				if (cbProp == sizeof(GUID))
				{
					LPTSTR szGUID = GUIDToString((LPGUID) lpCur);
					szData.FormatMessage(IDS_GUID);
					szData += _T(" = "); // STRING_OK
					szData += szGUID;
					bDataHandled = true;
				}
			}
			break;
		case EFPB_SFTAG:
			{
				DWORD swSFTag = NULL;
				if (cbProp == sizeof(DWORD)) 
				{
					swSFTag = *((DWORD*)lpCur);
					szData.Format(_T("0x%08X"),swSFTag); // STRING_OK
					bDataHandled = true;
				}
			}
			break;
		case EFPB_TODO_VERSION:
			{
				DWORD dwToDoVers = NULL;
				if (cbProp == 4) 
				{
					dwToDoVers = *((DWORD*)lpCur);
					szData.Format(_T("0x%08X"),dwToDoVers); // STRING_OK
					bDataHandled = true;
				}
			}
			break;
		}

		if (!bDataHandled)
		{
			CString szUnknownData;
			SBinary sBin = {0};
			sBin.cb = cbProp;
			sBin.lpb = lpCur;
			szData += BinToHexString(&sBin,true);
		}

		szResultString += szData;
		lpCur += cbProp;
		bFirstPass = false;
	}

	// junk routine
	if (lpCur < myBin.lpb + myBin.cb)
	{
		szResultString += _T("\n"); // STRING_OK
		SBinary sBin = {0};
		sBin.cb = (ULONG) (myBin.lpb + myBin.cb - lpCur);
		sBin.lpb = lpCur;
		szResultString += BinToHexString(&sBin,true);
	}

	size_t cchResultString = szResultString.GetLength()+1;
	szOut = new TCHAR[cchResultString];
	if (szOut)
	{
		HRESULT hRes = S_OK;
		EC_H(StringCchCopy(szOut,cchResultString,(LPCTSTR)szResultString));
	}
	*lpszResultString = szOut;
}

// Allocates return value with new.
// clean up with delete[].
TZDEFINITION* BinToTZDEFINITION(ULONG cbDef, LPBYTE lpbDef)
{
	if (!lpbDef) return NULL;

	// Update this if parsing code is changed!
	// this checks the size up to the flags member
	if (cbDef < 2*sizeof(BYTE) + 2*sizeof(WORD)) return NULL;

	TZDEFINITION tzDef = {0};
	TZRULE* lpRules = NULL;
	LPBYTE lpPtr = lpbDef;
	WORD cchKeyName = NULL;
	WCHAR* szKeyName = NULL;
	WORD i = 0;

	BYTE bMajorVersion = *((BYTE*)lpPtr);
	lpPtr += sizeof(BYTE);
	BYTE bMinorVersion = *((BYTE*)lpPtr);
	lpPtr += sizeof(BYTE);

	// We only understand TZ_BIN_VERSION_MAJOR
	if (TZ_BIN_VERSION_MAJOR != bMajorVersion) return NULL;

	// We only understand if >= TZ_BIN_VERSION_MINOR
	if (TZ_BIN_VERSION_MINOR > bMinorVersion) return NULL;

	lpPtr += sizeof(WORD);

	tzDef.wFlags = *((WORD*)lpPtr);
	lpPtr += sizeof(WORD);

	if (TZDEFINITION_FLAG_VALID_GUID & tzDef.wFlags)
	{
		if (lpbDef + cbDef < lpPtr + sizeof(GUID)) return NULL;
		tzDef.guidTZID = *((GUID*)lpPtr);
		lpPtr += sizeof(GUID);
	}

	if (TZDEFINITION_FLAG_VALID_KEYNAME & tzDef.wFlags)
	{
		if (lpbDef + cbDef < lpPtr + sizeof(WORD)) return NULL;
		cchKeyName = *((WORD*)lpPtr);
		lpPtr += sizeof(WORD);
		if (cchKeyName)
		{
			if (lpbDef + cbDef < lpPtr + (BYTE)sizeof(WORD)*cchKeyName) return NULL;
			szKeyName = (WCHAR*)lpPtr;
			lpPtr += cchKeyName*sizeof(WORD);
		}
	}

	if (lpbDef+ cbDef < lpPtr + sizeof(WORD)) return NULL;
	tzDef.cRules = *((WORD*)lpPtr);
	lpPtr += sizeof(WORD);

	if (tzDef.cRules)
	{
		lpRules = new TZRULE[tzDef.cRules];
		if (!lpRules) return NULL;

		LPBYTE lpNextRule = lpPtr;
		BOOL bRuleOK = false;

		for (i = 0;i<tzDef.cRules;i++)
		{
			bRuleOK = false;
			lpPtr = lpNextRule;

			if (lpbDef + cbDef < lpPtr + 2*sizeof(BYTE) + 2*sizeof(WORD) + 3*sizeof(long) + 2*sizeof(SYSTEMTIME)) return NULL;
			bRuleOK = true;
			BYTE bRuleMajorVersion = *((BYTE*)lpPtr);
			lpPtr += sizeof(BYTE);
			BYTE bRuleMinorVersion = *((BYTE*)lpPtr);
			lpPtr += sizeof(BYTE);

			// We only understand TZ_BIN_VERSION_MAJOR
			if (TZ_BIN_VERSION_MAJOR != bRuleMajorVersion) return NULL;

			// We only understand if >= TZ_BIN_VERSION_MINOR
			if (TZ_BIN_VERSION_MINOR > bRuleMinorVersion) return NULL;

			WORD cbRule = *((WORD*)lpPtr);
			lpPtr += sizeof(WORD);

			lpNextRule = lpPtr + cbRule;

			lpRules[i].wFlags = *((WORD*)lpPtr);
			lpPtr += sizeof(WORD);

			lpRules[i].stStart = *((SYSTEMTIME*)lpPtr);
			lpPtr += sizeof(SYSTEMTIME);

			lpRules[i].TZReg.lBias = *((long*)lpPtr);
			lpPtr += sizeof(long);
			lpRules[i].TZReg.lStandardBias = *((long*)lpPtr);
			lpPtr += sizeof(long);
			lpRules[i].TZReg.lDaylightBias = *((long*)lpPtr);
			lpPtr += sizeof(long);

			lpRules[i].TZReg.stStandardDate = *((SYSTEMTIME*)lpPtr);
			lpPtr += sizeof(SYSTEMTIME);
			lpRules[i].TZReg.stDaylightDate = *((SYSTEMTIME*)lpPtr);
			lpPtr += sizeof(SYSTEMTIME);
		}
		if (!bRuleOK)
		{
			delete[] lpRules;
			return NULL;
		}
	}

	// Now we've read everything - allocate a structure and copy it in

	size_t cbTZDef = sizeof(TZDEFINITION) +
		sizeof(WCHAR)*(cchKeyName+1) +
		sizeof(TZRULE)*tzDef.cRules;

	TZDEFINITION* ptzDef = (TZDEFINITION*) new BYTE[cbTZDef];

	if (ptzDef)
	{
		// Copy main struct over
		*ptzDef = tzDef;
		lpPtr = (LPBYTE) ptzDef;
		lpPtr += sizeof(TZDEFINITION);

		if (szKeyName)
		{
			ptzDef->pwszKeyName = (WCHAR*)lpPtr;
			memcpy(lpPtr,szKeyName,cchKeyName*sizeof(WCHAR));
			ptzDef->pwszKeyName[cchKeyName] = 0;

			lpPtr += (cchKeyName+1)*sizeof(WCHAR);
		}

		if (ptzDef->cRules && lpRules)
		{
			ptzDef->rgRules = (TZRULE*)lpPtr;
			for (i = 0;i<ptzDef->cRules;i++)
			{
				ptzDef->rgRules[i] = lpRules[i];
			}
		}
	}

	delete[] lpRules;

	return ptzDef;
}

// result allocated with new
// clean up with delete[]
LPTSTR TZDEFINITIONToString(TZDEFINITION tzDef)
{
	CString szDef;
	LPTSTR szOut = NULL;
	LPTSTR szGUID = GUIDToString(&tzDef.guidTZID);
	LPTSTR szFlags = NULL;
	HRESULT hRes = S_OK;

	EC_H(InterpretFlags(flagTZDef, tzDef.wFlags, &szFlags));
	szDef.FormatMessage(IDS_TZDEFTOSTRING,
		tzDef.wFlags,
		szFlags,
		szGUID?szGUID:_T("null"),// STRING_OK
		tzDef.pwszKeyName,
		tzDef.cRules);
	delete[] szFlags;
	szFlags = NULL;

	if (tzDef.cRules && tzDef.rgRules)
	{
		CString szRule;

		if (szRule)
		{
			WORD i = 0;

			for (i = 0;i<tzDef.cRules;i++)
			{
				EC_H(InterpretFlags(flagTZRule, tzDef.rgRules[i].wFlags, &szFlags));
				szRule.FormatMessage(IDS_TZRULETOSTRING,
					i,
					tzDef.rgRules[i].wFlags,
					szFlags,
					tzDef.rgRules[i].stStart.wYear,
					tzDef.rgRules[i].stStart.wMonth,
					tzDef.rgRules[i].stStart.wDayOfWeek,
					tzDef.rgRules[i].stStart.wDay,
					tzDef.rgRules[i].stStart.wHour,
					tzDef.rgRules[i].stStart.wMinute,
					tzDef.rgRules[i].stStart.wSecond,
					tzDef.rgRules[i].stStart.wMilliseconds,
					tzDef.rgRules[i].TZReg.lBias,
					tzDef.rgRules[i].TZReg.lStandardBias,
					tzDef.rgRules[i].TZReg.lDaylightBias,
					tzDef.rgRules[i].TZReg.stStandardDate.wYear,
					tzDef.rgRules[i].TZReg.stStandardDate.wMonth,
					tzDef.rgRules[i].TZReg.stStandardDate.wDayOfWeek,
					tzDef.rgRules[i].TZReg.stStandardDate.wDay,
					tzDef.rgRules[i].TZReg.stStandardDate.wHour,
					tzDef.rgRules[i].TZReg.stStandardDate.wMinute,
					tzDef.rgRules[i].TZReg.stStandardDate.wSecond,
					tzDef.rgRules[i].TZReg.stStandardDate.wMilliseconds,
					tzDef.rgRules[i].TZReg.stDaylightDate.wYear,
					tzDef.rgRules[i].TZReg.stDaylightDate.wMonth,
					tzDef.rgRules[i].TZReg.stDaylightDate.wDayOfWeek,
					tzDef.rgRules[i].TZReg.stDaylightDate.wDay,
					tzDef.rgRules[i].TZReg.stDaylightDate.wHour,
					tzDef.rgRules[i].TZReg.stDaylightDate.wMinute,
					tzDef.rgRules[i].TZReg.stDaylightDate.wSecond,
					tzDef.rgRules[i].TZReg.stDaylightDate.wMilliseconds);
				delete[] szFlags;
				szFlags = NULL;
				szDef += szRule;
			}
		}
	}

	size_t cchDef = szDef.GetLength()+1;
	szOut = new TCHAR[cchDef];
	if (szOut)
	{
		HRESULT hRes = S_OK;
		EC_H(StringCchCopy(szOut,cchDef,(LPCTSTR)szDef));
	}

	delete[] szGUID;
	return szOut;
}

// Allocates return value with new.
// clean up with delete.
TZREG* BinToTZREG(ULONG cbReg, LPBYTE lpbReg)
{
	if (!lpbReg) return NULL;

	// Update this if parsing code is changed!
	if (cbReg < 3*sizeof(long) + 2*sizeof(WORD) + 2*sizeof(SYSTEMTIME)) return NULL;

	TZREG tzReg = {0};
	LPBYTE lpPtr = lpbReg;

	tzReg.lBias = *((long*)lpPtr);
	lpPtr += sizeof(long);
	tzReg.lStandardBias = *((long*)lpPtr);
	lpPtr += sizeof(long);
	tzReg.lDaylightBias = *((long*)lpPtr);
	lpPtr += sizeof(long);
	lpPtr += sizeof(WORD);// reserved

	tzReg.stStandardDate = *((SYSTEMTIME*)lpPtr);
	lpPtr += sizeof(SYSTEMTIME);
	lpPtr += sizeof(WORD);// reserved
	tzReg.stDaylightDate = *((SYSTEMTIME*)lpPtr);
	lpPtr += sizeof(SYSTEMTIME);

	TZREG* ptzReg = NULL;
	ptzReg = new TZREG;
	if (ptzReg)
	{
		*ptzReg = tzReg;
	}

	return ptzReg;
}

// result allocated with new
// clean up with delete[]
LPTSTR TZREGToString(TZREG tzReg)
{
	HRESULT hRes = S_OK;
	CString szReg;
	LPTSTR szOut = NULL;

	szReg.FormatMessage(IDS_TZREGTOSTRING,
		tzReg.lBias,
		tzReg.lStandardBias,
		tzReg.lDaylightBias,
		tzReg.stStandardDate.wYear,
		tzReg.stStandardDate.wMonth,
		tzReg.stStandardDate.wDayOfWeek,
		tzReg.stStandardDate.wDay,
		tzReg.stStandardDate.wHour,
		tzReg.stStandardDate.wMinute,
		tzReg.stStandardDate.wSecond,
		tzReg.stStandardDate.wMilliseconds,
		tzReg.stDaylightDate.wYear,
		tzReg.stDaylightDate.wMonth,
		tzReg.stDaylightDate.wDayOfWeek,
		tzReg.stDaylightDate.wDay,
		tzReg.stDaylightDate.wHour,
		tzReg.stDaylightDate.wMinute,
		tzReg.stDaylightDate.wSecond,
		tzReg.stDaylightDate.wMilliseconds);

	size_t cchReg = szReg.GetLength()+1;
	szOut = new TCHAR[cchReg];
	if (szOut)
	{
		EC_H(StringCchCopy(szOut,cchReg,(LPCTSTR)szReg));
	}
	return szOut;
}
