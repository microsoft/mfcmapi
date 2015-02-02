#include "stdafx.h"
#include "InterpretProp2.h"
#include "InterpretProp.h"
#include "MAPIFunctions.h"
#include "NamedPropCache.h"
#include "String.h"
#include <vector>
#include <algorithm>
#include <unordered_map>

#define ulNoMatch 0xffffffff
static WCHAR szPropSeparator[] = L", "; // STRING_OK

// Compare tag sort order.
bool CompareTagsSortOrder(int a1, int a2)
{
	LPNAME_ARRAY_ENTRY_V2 lpTag1 = &PropTagArray[a1];
	LPNAME_ARRAY_ENTRY_V2 lpTag2 = &PropTagArray[a2];;

	if (lpTag1->ulSortOrder < lpTag2->ulSortOrder) return false;
	if (lpTag1->ulSortOrder == lpTag2->ulSortOrder)
	{
		return wcscmp(lpTag1->lpszName, lpTag2->lpszName)>0 ? false : true;
	}
	return true;
}

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
	std::vector<ULONG>& ulExacts,
	std::vector<ULONG>& ulPartials)
{
	if (!(ulTarget & PROP_TAG_MASK)) // not dealing with a full prop tag
	{
		ulTarget = PROP_TAG(PT_UNSPECIFIED, ulTarget);
	}

	ULONG ulLowerBound = 0;
	ULONG ulUpperBound = ulMyArray - 1; // ulMyArray-1 is the last entry
	ULONG ulMidPoint = (ulUpperBound + ulLowerBound) / 2;
	ULONG ulFirstMatch = ulNoMatch;
	ULONG ulMaskedTarget = ulTarget & PROP_TAG_MASK;

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
		else if (ulMaskedTarget >(PROP_TAG_MASK & MyArray[ulMidPoint].ulValue))
		{
			ulLowerBound = ulMidPoint;
		}

		ulMidPoint = (ulUpperBound + ulLowerBound) / 2;
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
		// Scan backwards to find the first partial match
		while (ulFirstMatch > 0 && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulFirstMatch - 1].ulValue))
		{
			ulFirstMatch = ulFirstMatch - 1;
		}

		// Grab our matches
		ULONG ulCur;
		for (ulCur = ulFirstMatch; ulCur < ulMyArray && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulCur].ulValue); ulCur++)
		{
			if (ulTarget == MyArray[ulCur].ulValue)
			{
				ulExacts.push_back(ulCur);
			}
			else
			{
				ulPartials.push_back(ulCur);
			}
		}

		if (ulExacts.size()) std::sort(ulExacts.begin(), ulExacts.end(), CompareTagsSortOrder);
		if (ulPartials.size()) std::sort(ulPartials.begin(), ulPartials.end(), CompareTagsSortOrder);

	}
}

struct NameMapEntry
{
	std::wstring szExactMatch;
	std::wstring szPartialMatches;
};

std::unordered_map<ULONG64, NameMapEntry> g_PropNames;

// lpszExactMatch and lpszPartialMatches allocated with new
// clean up with delete[]
void PropTagToPropName(ULONG ulPropTag, bool bIsAB, _Deref_opt_out_opt_z_ LPTSTR* lpszExactMatch, _Deref_opt_out_opt_z_ LPTSTR* lpszPartialMatches)
{
	if (lpszExactMatch) *lpszExactMatch = NULL;
	if (lpszPartialMatches) *lpszPartialMatches = NULL;
	if (!lpszExactMatch && !lpszPartialMatches) return;

	ULONG64 ulKey = (bIsAB ? (_int64)1 << 32 : 0) | ulPropTag;

	auto match = g_PropNames.find(ulKey);
	if (match != g_PropNames.end())
	{
		if (lpszExactMatch)
		{
			*lpszExactMatch = wstringToLPTSTR(match->second.szExactMatch);
		}

		if (lpszPartialMatches)
		{
			*lpszPartialMatches = wstringToLPTSTR(match->second.szPartialMatches);
		}

		return;
	}

	std::vector<ULONG> ulExacts;
	std::vector<ULONG> ulPartials;
	FindTagArrayMatches(ulPropTag, bIsAB, PropTagArray, ulPropTagArray, ulExacts, ulPartials);

	NameMapEntry entry;

	if (lpszExactMatch)
	{
		if (ulExacts.size())
		{
			for (ULONG ulMatch : ulExacts)
			{
				entry.szExactMatch += format(L"%ws", PropTagArray[ulMatch].lpszName);
				if (ulMatch != ulExacts.back())
				{
					entry.szExactMatch += szPropSeparator;
				}
			}

			*lpszExactMatch = wstringToLPTSTR(entry.szExactMatch);
		}
	}

	if (lpszPartialMatches)
	{
		if (ulPartials.size())
		{
			{
				for (ULONG ulMatch : ulPartials)
				{
					entry.szPartialMatches += format(L"%ws", PropTagArray[ulMatch].lpszName);
					if (ulMatch != ulPartials.back())
					{
						entry.szPartialMatches += szPropSeparator;
					}
				}

				*lpszPartialMatches = wstringToLPTSTR(entry.szPartialMatches);
			}
		}
	}

	g_PropNames.insert({ ulKey, entry });
}

// Strictly does a lookup in the array. Does not convert otherwise
_Check_return_ HRESULT LookupPropName(_In_z_ LPCWSTR lpszPropName, _Out_ ULONG* ulPropTag)
{
	if (!lpszPropName || !ulPropTag) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	ULONG ulCur = 0;

	*ulPropTag = NULL;

	if (!ulPropTagArray || !PropTagArray) return S_OK;

	for (ulCur = 0; ulCur < ulPropTagArray; ulCur++)
	{
		if (0 == lstrcmpiW(lpszPropName, PropTagArray[ulCur].lpszName))
		{
			*ulPropTag = PropTagArray[ulCur].ulValue;
			break;
		}
	}

	return hRes;
} // LookupPropName

_Check_return_ HRESULT PropNameToPropTagW(_In_z_ LPCWSTR lpszPropName, _Out_ ULONG* ulPropTag)
{
	if (!lpszPropName || !ulPropTag) return MAPI_E_INVALID_PARAMETER;

	*ulPropTag = NULL;

	LPWSTR szEnd = NULL;
	ULONG ulTag = wcstoul(lpszPropName, &szEnd, 16);
	if (*szEnd == NULL)
	{
		*ulPropTag = ulTag;
		return S_OK;
	}

	return LookupPropName(lpszPropName, ulPropTag);;
} // PropNameToPropTagW

_Check_return_ HRESULT PropNameToPropTagA(_In_z_ LPCSTR lpszPropName, _Out_ ULONG* ulPropTag)
{
	if (!lpszPropName || !ulPropTag) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	*ulPropTag = NULL;
	if (!ulPropTagArray || !PropTagArray) return S_OK;

	LPWSTR szPropName = NULL;
	EC_H(AnsiToUnicode(lpszPropName, &szPropName));
	if (SUCCEEDED(hRes))
	{
		EC_H(PropNameToPropTagW(szPropName, ulPropTag));
	}
	delete[] szPropName;
	return hRes;
} // PropNameToPropTagA

_Check_return_ ULONG PropTypeNameToPropTypeA(_In_z_ LPCSTR lpszPropType)
{
	ULONG ulPropType = PT_UNSPECIFIED;

	HRESULT hRes = S_OK;
	LPWSTR szPropType = NULL;
	EC_H(AnsiToUnicode(lpszPropType, &szPropType));
	ulPropType = PropTypeNameToPropTypeW(szPropType);
	delete[] szPropType;

	return ulPropType;
} // PropTypeNameToPropTypeA

_Check_return_ ULONG PropTypeNameToPropTypeW(_In_z_ LPCWSTR lpszPropType)
{
	if (!lpszPropType || !ulPropTypeArray || !PropTypeArray) return PT_UNSPECIFIED;

	// Check for numbers first before trying the string as an array lookup.
	// This will translate '0x102' to 0x102, 0x3 to 3, etc.
	LPWSTR szEnd = NULL;
	ULONG ulType = wcstoul(lpszPropType, &szEnd, 16);
	if (*szEnd == NULL) return ulType;

	ULONG ulCur = 0;

	ULONG ulPropType = PT_UNSPECIFIED;

	LPCWSTR szPropType = lpszPropType;
	for (ulCur = 0; ulCur < ulPropTypeArray; ulCur++)
	{
		if (0 == lstrcmpiW(szPropType, PropTypeArray[ulCur].lpszName))
		{
			ulPropType = PropTypeArray[ulCur].ulValue;
			break;
		}
	}

	return ulPropType;
} // PropTypeNameToPropTypeW

_Check_return_ wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID)
{
	ULONG ulCur = 0;
	wstring szGUID = GUIDToString(lpGUID);

	szGUID += L" = "; // STRING_OK

	if (lpGUID && ulPropGuidArray && PropGuidArray)
	{
		for (ulCur = 0; ulCur < ulPropGuidArray; ulCur++)
		{
			if (IsEqualGUID(*lpGUID, *PropGuidArray[ulCur].lpGuid))
			{
				szGUID += PropGuidArray[ulCur].lpszName;
				return szGUID;
			}
		}
	}

	return szGUID + formatmessage(IDS_UNKNOWNGUID);
}

LPCGUID GUIDNameToGUIDInt(_In_z_ LPCTSTR szGUID, bool bByteSwapped)
{
	HRESULT hRes = S_OK;
	LPGUID lpGuidRet = NULL;
	LPCGUID lpGUID = NULL;
	GUID guid = { 0 };

	// Try the GUID like PS_* first
	if (ulPropGuidArray && PropGuidArray)
	{
		ULONG ulCur = 0;
#ifdef UNICODE
		LPCWSTR szGUIDW = szGUID;
#else
		LPWSTR szGUIDW = NULL;
		EC_H(AnsiToUnicode(szGUID, &szGUIDW));
		if (SUCCEEDED(hRes))
		{
#endif
			for (ulCur = 0; ulCur < ulPropGuidArray; ulCur++)
			{
				if (0 == lstrcmpiW(szGUIDW, PropGuidArray[ulCur].lpszName))
				{
					lpGUID = PropGuidArray[ulCur].lpGuid;
					break;
				}
			}
#ifndef UNICODE
		}
		delete[] szGUIDW;
#endif
	}

	if (!lpGUID) // no match - try it like a guid {}
	{
		hRes = S_OK;
		WC_H(StringToGUID(szGUID, bByteSwapped, &guid));

		if (SUCCEEDED(hRes))
		{
			lpGUID = &guid;
		}
	}

	if (lpGUID)
	{
		lpGuidRet = new GUID;
		if (lpGuidRet)
		{
			memcpy(lpGuidRet, lpGUID, sizeof(GUID));
		}
	}

	return lpGuidRet;
}

LPCGUID GUIDNameToGUIDW(_In_z_ LPCWSTR szGUID, bool bByteSwapped)
{
#ifdef UNICODE
	return GUIDNameToGUIDInt(szGUID, bByteSwapped);
#else
	LPCGUID lpGUID = NULL;
	LPSTR szGUIDA = NULL;
	HRESULT hRes = UnicodeToAnsi(szGUID, &szGUIDA);
	if (SUCCEEDED(hRes) && szGUIDA)
	{
		lpGUID = GUIDNameToGUIDInt(szGUIDA, bByteSwapped);
		delete[] szGUIDA;
	}

	return lpGUID;
#endif
}

LPCGUID GUIDNameToGUIDA(_In_z_ LPCSTR szGUID, bool bByteSwapped)
{
#ifdef UNICODE
	LPCGUID lpGUID = NULL;
	LPWSTR szGUIDW = NULL;
	(void)AnsiToUnicode(szGUID, &szGUIDW);
	if (szGUIDW)
	{
		lpGUID = GUIDNameToGUIDInt(szGUIDW, bByteSwapped);
	}

	delete[] szGUIDW;

	return lpGUID;
#else
	return GUIDNameToGUIDInt(szGUID, bByteSwapped);
#endif
}

// Returns string built from NameIDArray
std::wstring NameIDToPropName(_In_ LPMAPINAMEID lpNameID)
{
	std::wstring szResultString;
	if (!lpNameID) return szResultString;
	if (lpNameID->ulKind != MNID_ID) return szResultString;
	ULONG ulCur = 0;
	ULONG ulMatch = ulNoMatch;

	if (!ulNameIDArray || !NameIDArray) return szResultString;

	for (ulCur = 0; ulCur < ulNameIDArray; ulCur++)
	{
		if (NameIDArray[ulCur].lValue == lpNameID->Kind.lID)
		{
			ulMatch = ulCur;
			break;
		}
	}

	if (ulNoMatch != ulMatch)
	{
		for (ulCur = ulMatch; ulCur < ulNameIDArray; ulCur++)
		{
			if (NameIDArray[ulCur].lValue != lpNameID->Kind.lID) break;
			// We don't acknowledge array entries without guids
			if (!NameIDArray[ulCur].lpGuid) continue;
			// But if we weren't asked about a guid, we don't check one
			if (lpNameID->lpguid && !IsEqualGUID(*lpNameID->lpguid, *NameIDArray[ulCur].lpGuid)) continue;

			if (ulCur != ulMatch)
			{
				szResultString += szPropSeparator;
			}

			szResultString += NameIDArray[ulCur].lpszName;
		}
	}

	return szResultString;
}

// Interprets a flag value according to a flag name and returns a string
// allocated with new
// Free the string with delete[]
// Will not return a string if the flag name is not recognized
void InterpretFlags(const __NonPropFlag ulFlagName, const LONG lFlagValue, _Deref_out_opt_z_ LPTSTR* szFlagString)
{
	InterpretFlags(ulFlagName, lFlagValue, _T(""), szFlagString);
} // InterpretFlags

_Check_return_  wstring InterpretFlags(const __NonPropFlag ulFlagName, const LONG lFlagValue)
{
	LPTSTR szFlags = NULL;
	InterpretFlags(ulFlagName, lFlagValue, &szFlags);
	wstring szRet = LPTSTRToWstring(szFlags);

	delete[] szFlags;
	return szRet;
}

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

	for (; FlagArray[ulCurEntry].ulFlagName == ulFlagName; ulCurEntry++)
	{
		if (flagFLAG == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue & lTempValue)
			{
				if (bNeedSeparator)
				{
					EC_H(StringCchCatW(szTempString, _countof(szTempString), L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString, _countof(szTempString), FlagArray[ulCurEntry].lpszName));
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
					EC_H(StringCchCatW(szTempString, _countof(szTempString), L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString, _countof(szTempString), FlagArray[ulCurEntry].lpszName));
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
					EC_H(StringCchCatW(szTempString, _countof(szTempString), L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString, _countof(szTempString), FlagArray[ulCurEntry].lpszName));
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
					EC_H(StringCchCatW(szTempString, _countof(szTempString), L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString, _countof(szTempString), FlagArray[ulCurEntry].lpszName));
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
					EC_H(StringCchCatW(szTempString, _countof(szTempString), L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString, _countof(szTempString), FlagArray[ulCurEntry].lpszName));
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
					EC_H(StringCchCatW(szTempString, _countof(szTempString), L" | ")); // STRING_OK
				}
				EC_H(StringCchCatW(szTempString, _countof(szTempString), FlagArray[ulCurEntry].lpszName));
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
					EC_H(StringCchCatW(szTempString, _countof(szTempString), L" | ")); // STRING_OK
				}
				WCHAR szClearedBits[15];
				EC_H(StringCchPrintfW(szClearedBits, _countof(szClearedBits), L"0x%X", lClearedBits)); // STRING_OK
				EC_H(StringCchCatW(szTempString, _countof(szTempString), szClearedBits));
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
			EC_H(StringCchCatW(szTempString, _countof(szTempString), L" | ")); // STRING_OK
		}
		EC_H(StringCchPrintfW(szUnk, _countof(szUnk), L"0x%X", lTempValue)); // STRING_OK
		EC_H(StringCchCatW(szTempString, _countof(szTempString), szUnk));
	}

	// Copy the string we computed for output
	size_t cchLen = 0;
	EC_H(StringCchLengthW(szTempString, _countof(szTempString), &cchLen));

	if (cchLen)
	{
		cchLen++; // for the NULL
		size_t cchPrefix = NULL;
		if (szPrefix)
		{
			EC_H(StringCchLength(szPrefix, STRSAFE_MAX_CCH, &cchPrefix));
			cchLen += cchPrefix;
		}

		*szFlagString = new TCHAR[cchLen];

		if (*szFlagString)
		{
			(*szFlagString)[0] = NULL;
			EC_H(StringCchPrintf(*szFlagString, cchLen, _T("%s%ws"), szPrefix ? szPrefix : _T(""), szTempString)); // STRING_OK
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
	for (; FlagArray[ulCurEntry].ulFlagName == ulFlagName; ulCurEntry++)
	{
		if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
		{
			// keep going
		}
		else
		{
			if (bHex)
			{
				szTempString.FormatMessage(IDS_FLAGTOSTRINGHEX, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName);
			}
			else
			{
				szTempString.FormatMessage(IDS_FLAGTOSTRINGDEC, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName);
			}
			szFlagString += szTempString;
		}
	}

	return szFlagString;
} // AllFlagsToString

// Returns LPSPropValue with value of a property
// Uses GetProps and falls back to OpenProperty if the value is large
// Free with MAPIFreeBuffer
_Check_return_ HRESULT GetLargeProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp)
{
	if (!lpMAPIProp || !lppProp) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric, _T("GetLargeProp getting buffer from 0x%08X\n"), ulPropTag);

	HRESULT			hRes = S_OK;
	ULONG			cValues = 0;
	LPSPropValue	lpPropArray = NULL;
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
		DebugPrint(DBGGeneric, _T("GetLargeProp property reported in GetProps as large.\n"));
		MAPIFreeBuffer(lpPropArray);
		lpPropArray = NULL;
		// need to get the data as a stream
		LPSTREAM lpStream = NULL;

		WC_MAPI(lpMAPIProp->OpenProperty(
			ulPropTag,
			&IID_IStream,
			STGM_READ,
			0,
			(LPUNKNOWN*)&lpStream));
		if (SUCCEEDED(hRes) && lpStream)
		{
			STATSTG	StatInfo = { 0 };
			lpStream->Stat(&StatInfo, STATFLAG_NONAME); // find out how much space we need

			// We're not going to try to support MASSIVE properties.
			if (!StatInfo.cbSize.HighPart)
			{
				EC_H(MAPIAllocateBuffer(
					sizeof(SPropValue),
					(LPVOID*)&lpPropArray));
				if (lpPropArray)
				{
					memset(lpPropArray, 0, sizeof(SPropValue));
					lpPropArray->ulPropTag = ulPropTag;

					if (StatInfo.cbSize.LowPart)
					{
						EC_H(MAPIAllocateMore(
							StatInfo.cbSize.LowPart,
							lpPropArray,
							(LPVOID*)&lpPropArray->Value.bin.lpb));
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
		DebugPrint(DBGGeneric, _T("GetLargeProp GetProps found property.\n"));
		bSuccess = true;
	}
	else if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag))
	{
		DebugPrint(DBGGeneric, _T("GetLargeProp GetProps reported property as error 0x%08X.\n"), lpPropArray->Value.err);
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

// Returns LPSPropValue with value of a binary property
// Free with MAPIFreeBuffer
_Check_return_ HRESULT GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp)
{
	return GetLargeProp(lpMAPIProp, CHANGE_PROP_TYPE(ulPropTag, PT_BINARY), lppProp);
}

// Returns LPSPropValue with value of a string property
// Free with MAPIFreeBuffer
_Check_return_ HRESULT GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp)
{
	return GetLargeProp(lpMAPIProp, CHANGE_PROP_TYPE(ulPropTag, PT_TSTRING), lppProp);
}