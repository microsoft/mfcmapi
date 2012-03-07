#include "stdafx.h"
#include "InterpretProp2.h"
#include "InterpretProp.h"
#include "MAPIFunctions.h"
#include "NamedPropCache.h"

#define ulNoMatch 0xffffffff
static WCHAR szPropSeparator[] = L", "; // STRING_OK

// Searches an array for a target number.
// Search is done with a mask
// Partial matches are those that match with the mask applied
// Exact matches are those that match without the mask applied
// lpUlNumPartials will exclude count of exact matches
// if it wants just the true partial matches.
// If no hits, then ulNoMatch should be returned for lpulFirstExact and/or lpulFirstPartial
void FindTagArrayMatches(_In_ ULONG ulTarget,
						 bool bIsAB,
						 _In_count_(ulMyArray) NAME_ARRAY_ENTRY_V2* MyArray,
						 _In_ ULONG ulMyArray,
						 _Out_ ULONG* lpulNumExacts,
						 _Out_ ULONG* lpulFirstExact,
						 _Out_ ULONG* lpulNumPartials,
						 _Out_ ULONG* lpulFirstPartial)
{
	if (!(ulTarget & PROP_TAG_MASK)) // not dealing with a full prop tag
	{
		ulTarget = PROP_TAG(PT_UNSPECIFIED,ulTarget);
	}

	ULONG ulLowerBound = 0;
	ULONG ulUpperBound = ulMyArray-1; // ulMyArray-1 is the last entry
	ULONG ulMidPoint = (ulUpperBound+ulLowerBound)/2;
	ULONG ulFirstMatch = ulNoMatch;
	ULONG ulLastMatch = ulNoMatch;
	ULONG ulFirstExactMatch = ulNoMatch;
	ULONG ulLastExactMatch = ulNoMatch;
	ULONG ulMaskedTarget = ulTarget & PROP_TAG_MASK;

	if (lpulNumExacts) *lpulNumExacts = 0;
	if (lpulFirstExact) *lpulFirstExact = ulNoMatch;
	if (lpulNumPartials) *lpulNumPartials = 0;
	if (lpulFirstPartial) *lpulFirstPartial = ulNoMatch;

	// Short circuit property IDs with the high bit set if bIsAB wasn't passed
	if (!bIsAB && (ulTarget & 0x80000000)) return;

	// Find A partial match
	while (ulUpperBound - ulLowerBound > 1)
	{
		if (ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulMidPoint].ulValue))
		{
			ulFirstMatch = ulMidPoint;
			break;
		}

		if (ulMaskedTarget < (PROP_TAG_MASK & MyArray[ulMidPoint].ulValue))
		{
			ulUpperBound = ulMidPoint;
		}
		else if (ulMaskedTarget > (PROP_TAG_MASK & MyArray[ulMidPoint].ulValue))
		{
			ulLowerBound = ulMidPoint;
		}
		ulMidPoint = (ulUpperBound+ulLowerBound)/2;
	}

	// When we get down to two points, we may have only checked one of them
	// Make sure we've checked the other
	if (ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulUpperBound].ulValue))
	{
		ulFirstMatch = ulUpperBound;
	}
	else if (ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulLowerBound].ulValue))
	{
		ulFirstMatch = ulLowerBound;
	}

	// Check that we got a match
	if (ulNoMatch != ulFirstMatch)
	{
		ulLastMatch = ulFirstMatch; // Remember the last match we've found so far

		// Scan backwards to find the first partial match
		while (ulFirstMatch > 0 && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulFirstMatch-1].ulValue))
		{
			ulFirstMatch = ulFirstMatch - 1;
		}

		// Scan forwards to find the real last partial match
		// Last entry in the array is ulMyArray-1
		while (ulLastMatch+1 < ulMyArray && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulLastMatch+1].ulValue))
		{
			ulLastMatch = ulLastMatch + 1;
		}

		// Scan to see if we have any exact matches
		ULONG ulCur;
		for (ulCur = ulFirstMatch ; ulCur <= ulLastMatch ; ulCur++)
		{
			if (ulTarget == MyArray[ulCur].ulValue)
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
				if (ulTarget == MyArray[ulCur].ulValue)
				{
					ulLastExactMatch = ulCur;
				}
				else break;
			}
			ulNumExacts = ulLastExactMatch - ulFirstExactMatch + 1;
		}

		ULONG ulNumPartials = ulLastMatch - ulFirstMatch + 1 - ulNumExacts;

		if (lpulNumExacts) *lpulNumExacts = ulNumExacts;
		if (lpulFirstExact) *lpulFirstExact = ulFirstExactMatch;
		if (lpulNumPartials) *lpulNumPartials = ulNumPartials;
		if (lpulFirstPartial) *lpulFirstPartial = ulFirstMatch;
	}
} // FindTagArrayMatches

// Compare tag sort order. 
int _cdecl CompareTagsSortOrder(_In_ const void* a1, _In_ const void* a2)
{
	LPNAME_ARRAY_ENTRY_V2 lpTag1 = &PropTagArray[* (LPULONG) a1];
	LPNAME_ARRAY_ENTRY_V2 lpTag2 = &PropTagArray[* (LPULONG) a2];;

	if (lpTag1->ulSortOrder < lpTag2->ulSortOrder) return 1;
	if (lpTag1->ulSortOrder == lpTag2->ulSortOrder)
	{
		return wcscmp(lpTag1->lpszName,lpTag2->lpszName);
	}
	return -1;
} // CompareTagsSortOrder

// lpszExactMatch and lpszPartialMatches allocated with new
// clean up with delete[]
// The compiler gets confused by the qsort call and thinks lpulPartials is smaller than it really is.
// Then it complains when I write to the 'bad' memory, even though it's definitely good.
// This is a bug in SAL, so I'm disabling the warning.
#pragma warning(push)
#pragma warning(disable:6385)
_Check_return_ HRESULT PropTagToPropName(ULONG ulPropTag, bool bIsAB, _Deref_opt_out_opt_z_ LPTSTR* lpszExactMatch, _Deref_opt_out_opt_z_ LPTSTR* lpszPartialMatches)
{
	if (!lpszExactMatch && !lpszPartialMatches) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	ULONG ulNumExacts = NULL;
	ULONG ulFirstExactMatch = ulNoMatch;
	ULONG ulNumPartials = NULL;
	ULONG ulFirstPartial = ulNoMatch;
	ULONG ulCur = NULL;
	ULONG i = 0;

	FindTagArrayMatches(ulPropTag,bIsAB,PropTagArray,ulPropTagArray,&ulNumExacts,&ulFirstExactMatch,&ulNumPartials,&ulFirstPartial);

	if (lpszExactMatch && ulNumExacts > 0 && ulNoMatch != ulFirstExactMatch)
	{
		ULONG ulLastExactMatch = ulFirstExactMatch+ulNumExacts-1;
		ULONG* lpulExacts = new ULONG[ulNumExacts];
		if (lpulExacts)
		{
			memset(lpulExacts,0,ulNumExacts*sizeof(ULONG));
			size_t cchExact = 1 + (ulNumExacts - 1) * (_countof(szPropSeparator) - 1);
			for (ulCur = ulFirstExactMatch ; ulCur <= ulLastExactMatch ; ulCur++)
			{
				size_t cchLen = 0;
				EC_H(StringCchLengthW(PropTagArray[ulCur].lpszName,STRSAFE_MAX_CCH,&cchLen));
				cchExact += cchLen;
				if (i < ulNumExacts) lpulExacts[i] = ulCur;
				i++;
			}

			qsort(lpulExacts,i,sizeof(ULONG),&CompareTagsSortOrder);

			LPWSTR szExactMatch = new WCHAR[cchExact];
			if (szExactMatch)
			{
				szExactMatch[0] = _T('\0');
				for (i = 0 ; i < ulNumExacts ; i++)
				{
					EC_H(StringCchCatW(szExactMatch,cchExact,PropTagArray[lpulExacts[i]].lpszName));
					if (i+1 < ulNumExacts)
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
		delete[] lpulExacts;
	}

	if (lpszPartialMatches && ulNumPartials > 0 && lpszPartialMatches)
	{
		ULONG* lpulPartials = new ULONG[ulNumPartials];
		if (lpulPartials)
		{
			memset(lpulPartials,0,ulNumPartials*sizeof(ULONG));
			// let's build lpszPartialMatches
			// see how much space we need - initialize cchPartial with space for separators and NULL terminator
			// note - ulNumPartials-1 is the number of spaces we need...
			ULONG ulLastMatch = ulFirstPartial+ulNumPartials+ulNumExacts-1;
			size_t cchPartial = 1 + (ulNumPartials - 1) * (_countof(szPropSeparator) - 1);
			i = 0;
			for (ulCur = ulFirstPartial ; ulCur <= ulLastMatch ; ulCur++)
			{
				if (ulPropTag == PropTagArray[ulCur].ulValue) continue; // skip our exact matches
				size_t cchLen = 0;
				EC_H(StringCchLengthW(PropTagArray[ulCur].lpszName,STRSAFE_MAX_CCH,&cchLen));
				cchPartial += cchLen;
				if (i < ulNumPartials) lpulPartials[i] = ulCur;
				i++;
			}

			qsort(lpulPartials,i,sizeof(ULONG),&CompareTagsSortOrder);

			LPWSTR szPartialMatches = new WCHAR[cchPartial];
			if (szPartialMatches)
			{
				szPartialMatches[0] = _T('\0');
				ULONG ulNumSeparators = 1; // start at 1 so we print one less than we print strings
				for (i = 0 ; i < ulNumPartials ; i++)
				{
					EC_H(StringCchCatW(szPartialMatches,cchPartial,PropTagArray[lpulPartials[i]].lpszName));
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
		delete[] lpulPartials;
	}

	return hRes;
} // PropTagToPropName
#pragma warning(pop)

_Check_return_ HRESULT PropNameToPropTagW(_In_z_ LPCWSTR lpszPropName, _Out_ ULONG* ulPropTag)
{
	if (!lpszPropName || !ulPropTag) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	ULONG ulCur = 0;

	*ulPropTag = NULL;
	if (!ulPropTagArray || !PropTagArray) return S_OK;
	LPCWSTR szPropName = lpszPropName;

	for (ulCur = 0 ; ulCur < ulPropTagArray ; ulCur++)
	{
		if (0 == lstrcmpiW(szPropName,PropTagArray[ulCur].lpszName))
		{
			*ulPropTag = PropTagArray[ulCur].ulValue;
			break;
		}
	}

	return hRes;
} // PropNameToPropTagW

_Check_return_ HRESULT PropNameToPropTagA(_In_z_ LPCSTR lpszPropName, _Out_ ULONG* ulPropTag)
{
	if (!lpszPropName || !ulPropTag) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	*ulPropTag = NULL;
	if (!ulPropTagArray || !PropTagArray) return S_OK;

	LPWSTR szPropName = NULL;
	EC_H(AnsiToUnicode(lpszPropName,&szPropName));
	if (SUCCEEDED(hRes))
	{
		EC_H(PropNameToPropTagW(szPropName,ulPropTag));
	}
	delete[] szPropName;
	return hRes;
} // PropNameToPropTagA

_Check_return_ ULONG PropTypeNameToPropTypeA(_In_z_ LPCSTR lpszPropType)
{
	ULONG ulPropType = PT_UNSPECIFIED;

	HRESULT hRes = S_OK;
	LPWSTR szPropType = NULL;
	EC_H(AnsiToUnicode(lpszPropType,&szPropType));
	ulPropType =  PropTypeNameToPropTypeW(szPropType);
	delete[] szPropType;

	return ulPropType;
} // PropTypeNameToPropTypeA

_Check_return_ ULONG PropTypeNameToPropTypeW(_In_z_ LPCWSTR lpszPropType)
{
	if (!lpszPropType || !ulPropTypeArray || !PropTypeArray) return PT_UNSPECIFIED;

	// Check for numbers first before trying the string as an array lookup.
	// This will translate '0x102' to 0x102, 0x3 to 3, etc.
	LPWSTR szEnd = NULL;
	ULONG ulType = wcstoul(lpszPropType,&szEnd,16);
	if (*szEnd == NULL) return ulType;

	ULONG ulCur = 0;

	ULONG ulPropType = PT_UNSPECIFIED;

	LPCWSTR szPropType = lpszPropType;
	for (ulCur = 0 ; ulCur < ulPropTypeArray ; ulCur++)
	{
		if (0 == lstrcmpiW(szPropType,PropTypeArray[ulCur].lpszName))
		{
			ulPropType = PropTypeArray[ulCur].ulValue;
			break;
		}
	}

	return ulPropType;
} // PropTypeNameToPropTypeW

_Check_return_ LPTSTR GUIDToStringAndName(_In_opt_ LPCGUID lpGUID)
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
			_countof(szUnknown)));
		szGUIDName = szUnknown;
	}

	LPTSTR szGUID = GUIDToString(lpGUID);

	EC_H(StringCchLengthW(szGUIDName,STRSAFE_MAX_CCH,&cchGUIDName));
	if (szGUID) EC_H(StringCchLength(szGUID,STRSAFE_MAX_CCH,&cchGUID));

	size_t cchBothGuid = cchGUID + 3 + cchGUIDName+1;
	LPTSTR szBothGuid = new TCHAR[cchBothGuid];

	if (szBothGuid)
	{
		EC_H(StringCchPrintf(szBothGuid,cchBothGuid,_T("%s = %ws"),szGUID,szGUIDName)); // STRING_OK
	}

	delete[] szGUID;
	return szBothGuid;
} // GUIDToStringAndName

void GUIDNameToGUID(_In_z_ LPCTSTR szGUID, _Deref_out_opt_ LPCGUID* lpGUID)
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
} // GUIDNameToGUID

// Allocates and returns string built from NameIDArray
// Allocated with new, clean up with delete[]
_Check_return_ LPWSTR NameIDToPropName(_In_ LPMAPINAMEID lpNameID)
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
		if (NameIDArray[ulCur].lValue != lpNameID->Kind.lID) break;
		// We don't acknowledge array entries without guids
		if (!NameIDArray[ulCur].lpGuid) continue;
		// But if we weren't asked about a guid, we don't check one
		if (lpNameID->lpguid && !IsEqualGUID(*lpNameID->lpguid,*NameIDArray[ulCur].lpGuid)) continue;

		EC_H(StringCchLengthW(NameIDArray[ulCur].lpszName,STRSAFE_MAX_CCH,&cchLen));
		cchResultString += cchLen;
		ulNumMatches++;
	}

	if (!ulNumMatches) return NULL;

	// Add in space for null terminator and separators
	cchResultString += 1 + (ulNumMatches - 1) * (_countof(szPropSeparator) - 1);

	// Copy our matches into the result string
	LPWSTR szResultString = NULL;
	szResultString = new WCHAR[cchResultString];
	if (szResultString)
	{
		szResultString[0] = L'\0'; // STRING_OK
		for (ulCur = ulMatch ; ulCur < ulNameIDArray ; ulCur++)
		{
			if (NameIDArray[ulCur].lValue != lpNameID->Kind.lID) break;
			// We don't acknowledge array entries without guids
			if (!NameIDArray[ulCur].lpGuid) continue;
			// But if we weren't asked about a guid, we don't check one
			if (lpNameID->lpguid && !IsEqualGUID(*lpNameID->lpguid,*NameIDArray[ulCur].lpGuid)) continue;

			EC_H(StringCchCatW(szResultString,cchResultString,NameIDArray[ulCur].lpszName));
			if (--ulNumMatches > 0)
			{
				EC_H(StringCchCatW(szResultString,cchResultString,szPropSeparator));
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
} // NameIDToPropName

// Interprets a flag value according to a flag name and returns a string
// allocated with new
// Free the string with delete[]
// Will not return a string if the flag name is not recognized
void InterpretFlags(const __NonPropFlag ulFlagName, const LONG lFlagValue, _Deref_out_opt_z_ LPTSTR* szFlagString)
{
	InterpretFlags(ulFlagName, lFlagValue, _T(""), szFlagString);
} // InterpretFlags

// Interprets a flag value according to a flag name and returns a string
// allocated with new
// Free the string with delete[]
// Will not return a string if the flag name is not recognized
void InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, _In_z_ LPCTSTR szPrefix, _Deref_out_opt_z_ LPTSTR* szFlagString)
{
	if (!szFlagString)
	{
		return;
	}
	HRESULT	hRes = S_OK;
	ULONG	ulCurEntry = 0;
	LONG	lTempValue = lFlagValue;
	WCHAR	szTempString[1024];

	*szFlagString = NULL;
	szTempString[0] = NULL;

	if (!ulFlagArray || !FlagArray) return;

	while (ulCurEntry < ulFlagArray && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
	{
		ulCurEntry++;
	}

	// Don't run off the end of the array
	if (ulFlagArray == ulCurEntry) return;
	if (FlagArray[ulCurEntry].ulFlagName != ulFlagName) return;

	// We've matched our flag name to the array - we SHOULD return a string at this point
	bool bNeedSeparator = false;

	for (;FlagArray[ulCurEntry].ulFlagName == ulFlagName;ulCurEntry++)
	{
		if (flagFLAG == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue & lTempValue)
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,_countof(szTempString),L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString,_countof(szTempString),FlagArray[ulCurEntry].lpszName));
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
					EC_H(StringCchCatW(szTempString,_countof(szTempString),L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString,_countof(szTempString),FlagArray[ulCurEntry].lpszName));
				lTempValue = 0;
				bNeedSeparator = true;
			}
		}
		else if (flagVALUEHIGHBYTES == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue == ((lTempValue >> 16) & 0xFFFF))
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,_countof(szTempString),L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString,_countof(szTempString),FlagArray[ulCurEntry].lpszName));
				lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 16);
				bNeedSeparator = true;
			}
		}
		else if (flagVALUE3RDBYTE == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue == ((lTempValue >> 8) & 0xFF))
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,_countof(szTempString),L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString,_countof(szTempString),FlagArray[ulCurEntry].lpszName));
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
					EC_H(StringCchCatW(szTempString,_countof(szTempString),L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString,_countof(szTempString),FlagArray[ulCurEntry].lpszName));
				lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
				bNeedSeparator = true;
			}
		}
		else if (flagVALUELOWERNIBBLE == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0x0F))
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,_countof(szTempString),L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString,_countof(szTempString),FlagArray[ulCurEntry].lpszName));
				lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
				bNeedSeparator = true;
			}
		}
		else if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
		{
			// find any bits we need to clear
			LONG lClearedBits = FlagArray[ulCurEntry].lFlagValue & lTempValue;
			// report what we found
			if (0 != lClearedBits)
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString,_countof(szTempString),L" | ")); // STRING_OK
				}
				WCHAR szClearedBits[15];
				EC_H(StringCchPrintfW(szClearedBits,_countof(szClearedBits),L"0x%X",lClearedBits)); // STRING_OK
				EC_H(StringCchCatW(szTempString,_countof(szTempString),szClearedBits));
				// clear the bits out
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
			EC_H(StringCchCatW(szTempString,_countof(szTempString),L" | ")); // STRING_OK
		}
		EC_H(StringCchPrintfW(szUnk,_countof(szUnk),L"0x%X",lTempValue)); // STRING_OK
		EC_H(StringCchCatW(szTempString,_countof(szTempString),szUnk));
	}

	// Copy the string we computed for output
	size_t cchLen = 0;
	EC_H(StringCchLengthW(szTempString,_countof(szTempString),&cchLen));

	if (cchLen)
	{
		cchLen++; // for the NULL
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
			EC_H(StringCchPrintf(*szFlagString,cchLen,_T("%s%ws"),szPrefix?szPrefix:_T(""),szTempString)); // STRING_OK
		}
	}
} // InterpretFlags

// Returns a list of all known flags/values for a flag name.
// For instance, for flagFuzzyLevel, would return:
// \r\n0x00000000 FL_FULLSTRING\r\n\
// 0x00000001 FL_SUBSTRING\r\n\
// 0x00000002 FL_PREFIX\r\n\
// 0x00010000 FL_IGNORECASE\r\n\
// 0x00020000 FL_IGNORENONSPACE\r\n\
// 0x00040000 FL_LOOSE
//
// Since the string is always appended to a prompt we include \r\n at the start
_Check_return_ CString AllFlagsToString(const ULONG ulFlagName, bool bHex)
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

	// We've matched our flag name to the array - we SHOULD return a string at this point
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
				szTempString.FormatMessage(IDS_FLAGTOSTRINGHEX,FlagArray[ulCurEntry].lFlagValue,FlagArray[ulCurEntry].lpszName);
			}
			else
			{
				szTempString.FormatMessage(IDS_FLAGTOSTRINGDEC,FlagArray[ulCurEntry].lFlagValue,FlagArray[ulCurEntry].lpszName);
			}
			szFlagString += szTempString;
		}
	}

	return szFlagString;
} // AllFlagsToString

// Uber property interpreter - given an LPSPropValue, produces all manner of strings
// lpszNameExactMatches, lpszNamePartialMatches allocated with new, delete with delete[]
// lpszNamedPropName, lpszNamedPropGUID, lpszNamedPropDASL freed with FreeNameIDStrings
// If lpProp is NULL but ulPropTag and lpMAPIProp are passed, will call GetProps
void InterpretProp(_In_opt_ LPSPropValue lpProp, // optional property value
				   ULONG ulPropTag, // optional 'original' prop tag
				   _In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
				   _In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
				   _In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
				   bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
				   _Deref_out_opt_z_ LPTSTR* lpszNameExactMatches, // Built from ulPropTag & bIsAB
				   _Deref_out_opt_z_ LPTSTR* lpszNamePartialMatches, // Built from ulPropTag & bIsAB
				   _In_opt_ CString* PropType, // Built from ulPropTag
				   _In_opt_ CString* PropTag, // Built from ulPropTag
				   _In_opt_ CString* PropString, // Built from lpProp
				   _In_opt_ CString* AltPropString, // Built from lpProp
				   _Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
				   _Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
				   _Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropDASL) // Built from ulPropTag & lpMAPIProp
{
	HRESULT hRes = S_OK;

	// These four strings are based on ulPropTag, not the LPSPropValue
	if (PropType) *PropType = TypeToString(ulPropTag);
	if (lpszNameExactMatches || lpszNamePartialMatches)
		EC_H(PropTagToPropName(ulPropTag,bIsAB,lpszNameExactMatches,lpszNamePartialMatches));
	if (PropTag) PropTag->Format(_T("0x%08X"),ulPropTag); // STRING_OK

	// Named Props
	LPMAPINAMEID* lppPropNames = 0;

	// If we weren't passed named property information and we need it, look it up
	// We don't check bIsAB here - some address book providers may actually support named props
	// If they don't, they should return empty results or an error, either of which we can handle
	if (!lpNameID &&
		lpMAPIProp && // if we have an object
		RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD && // and we're parsing named props
		(RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD || PROP_ID(ulPropTag) >= 0x8000) && // and it's either a named prop or we're doing all props
		(lpszNamedPropName || lpszNamedPropGUID || lpszNamedPropDASL)) // and we want to return something that needs named prop information
	{
		SPropTagArray	tag = {0};
		LPSPropTagArray	lpTag = &tag;
		ULONG			ulPropNames = 0;
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
	}

	// Avoid making the call if we don't have to so we don't accidently depend on MAPI
	if (lppPropNames) MAPIFreeBuffer(lppPropNames);
} // InterpretProp

// Returns LPSPropValue with value of a binary property
// Uses GetProps and falls back to OpenProperty if the value is large
// Free with MAPIFreeBuffer
_Check_return_ HRESULT GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp)
{
	if (!lpMAPIProp || !lppProp) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric,_T("GetLargeBinaryProp getting buffer from 0x%08X\n"),ulPropTag);

	ulPropTag = CHANGE_PROP_TYPE(ulPropTag,PT_BINARY);

	HRESULT			hRes		= S_OK;
	ULONG			cValues		= 0;
	LPSPropValue	lpPropArray	= NULL;
	bool			bSuccess = false;

	const SizedSPropTagArray(1, sptaBuffer) =
	{
		1,
		ulPropTag
	};
	*lppProp = NULL;

	WC_H_GETPROPS(lpMAPIProp->GetProps((LPSPropTagArray)&sptaBuffer, 0, &cValues, &lpPropArray));

	if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag) && MAPI_E_NOT_ENOUGH_MEMORY == lpPropArray->Value.err)
	{
		DebugPrint(DBGGeneric,_T("GetLargeBinaryProp property reported in GetProps as large.\n"));
		MAPIFreeBuffer(lpPropArray);
		lpPropArray = NULL;
		// need to get the data as a stream
		LPSTREAM lpStream = NULL;

		WC_MAPI(lpMAPIProp->OpenProperty(
			ulPropTag,
			&IID_IStream,
			STGM_READ,
			0,
			(LPUNKNOWN*) &lpStream));
		if (SUCCEEDED(hRes) && lpStream)
		{
			STATSTG	StatInfo = {0};
			lpStream->Stat(&StatInfo, STATFLAG_NONAME); // find out how much space we need

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
							EC_MAPI(lpStream->Read(lpPropArray->Value.bin.lpb, StatInfo.cbSize.LowPart, &lpPropArray->Value.bin.cb));
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
		DebugPrint(DBGGeneric,_T("GetLargeBinaryProp GetProps found property.\n"));
		bSuccess = true;
	}
	else if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag))
	{
		DebugPrint(DBGGeneric,_T("GetLargeBinaryProp GetProps reported property as error 0x%08X.\n"),lpPropArray->Value.err);
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
} // GetLargeBinaryProp