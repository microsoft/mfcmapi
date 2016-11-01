#include "stdafx.h"
#include "InterpretProp2.h"
#include "String.h"
#include <unordered_map>

#define ulNoMatch 0xffffffff
static WCHAR szPropSeparator[] = L", "; // STRING_OK

// Compare tag sort order.
bool CompareTagsSortOrder(int a1, int a2)
{
	auto lpTag1 = &PropTagArray[a1];
	auto lpTag2 = &PropTagArray[a2];;

	if (lpTag1->ulSortOrder < lpTag2->ulSortOrder) return false;
	if (lpTag1->ulSortOrder == lpTag2->ulSortOrder)
	{
		return wcscmp(lpTag1->lpszName, lpTag2->lpszName) > 0 ? false : true;
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
	const vector<NAME_ARRAY_ENTRY_V2>& MyArray,
	vector<ULONG>& ulExacts,
	vector<ULONG>& ulPartials)
{
	if (!(ulTarget & PROP_TAG_MASK)) // not dealing with a full prop tag
	{
		ulTarget = PROP_TAG(PT_UNSPECIFIED, ulTarget);
	}

	ULONG ulLowerBound = 0;
	auto ulUpperBound = static_cast<ULONG>(MyArray.size() - 1); // size-1 is the last entry
	auto ulMidPoint = (ulUpperBound + ulLowerBound) / 2;
	ULONG ulFirstMatch = ulNoMatch;
	auto ulMaskedTarget = ulTarget & PROP_TAG_MASK;

	// Short circuit property IDs with the high bit set if bIsAB wasn't passed
	if (!bIsAB && ulTarget & 0x80000000) return;

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
		for (ulCur = ulFirstMatch; ulCur < MyArray.size() && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulCur].ulValue); ulCur++)
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

		if (ulExacts.size()) sort(ulExacts.begin(), ulExacts.end(), CompareTagsSortOrder);
		if (ulPartials.size()) sort(ulPartials.begin(), ulPartials.end(), CompareTagsSortOrder);

	}
}

unordered_map<ULONG64, PropTagNames> g_PropNames;

PropTagNames PropTagToPropName(ULONG ulPropTag, bool bIsAB)
{
	auto ulKey = (bIsAB ? static_cast<ULONG64>(1) << 32 : 0) | ulPropTag;

	auto match = g_PropNames.find(ulKey);
	if (match != g_PropNames.end())
	{
		return match->second;
	}

	vector<ULONG> ulExacts;
	vector<ULONG> ulPartials;
	FindTagArrayMatches(ulPropTag, bIsAB, PropTagArray, ulExacts, ulPartials);

	PropTagNames entry;

	if (ulExacts.size())
	{
		for (auto ulMatch : ulExacts)
		{
			entry.exactMatches += PropTagArray[ulMatch].lpszName;
			if (ulMatch != ulExacts.back())
			{
				entry.exactMatches += szPropSeparator;
			}

			if (ulMatch == ulExacts.front())
			{
				entry.bestGuess = PropTagArray[ulMatch].lpszName;
			}
		}
	}

	if (ulPartials.size())
	{
		{
			for (auto ulMatch : ulPartials)
			{
				entry.partialMatches += PropTagArray[ulMatch].lpszName;
				if (ulMatch != ulPartials.back())
				{
					entry.partialMatches += szPropSeparator;
				}

				if (entry.bestGuess.empty() && ulMatch == ulPartials.front())
				{
					entry.bestGuess = PropTagArray[ulMatch].lpszName;
				}
			}
		}
	}

	// For PT_ERROR properties, we won't ever have an exact match
	// So we swap in all the partial matches instead
	if (PROP_TYPE(ulPropTag) == PT_ERROR && entry.exactMatches.empty())
	{
		entry.exactMatches.swap(entry.partialMatches);
	}

	g_PropNames.insert({ ulKey, entry });

	return entry;
}

// Strictly does a lookup in the array. Does not convert otherwise
_Check_return_ ULONG LookupPropName(_In_ const wstring& lpszPropName)
{
	if (lpszPropName.empty()) return 0;

	for (size_t ulCur = 0; ulCur < PropTagArray.size(); ulCur++)
	{
		if (0 == lstrcmpiW(lpszPropName.c_str(), PropTagArray[ulCur].lpszName))
		{
			return  PropTagArray[ulCur].ulValue;
		}
	}

	return 0;
}

_Check_return_ ULONG PropNameToPropTag(_In_ const wstring& lpszPropName)
{
	if (lpszPropName.empty()) return 0;

	auto ulTag = wstringToUlong(lpszPropName, 16);
	if (ulTag != NULL)
	{
		return ulTag;
	}

	return LookupPropName(lpszPropName);
}

_Check_return_ ULONG PropTypeNameToPropType(_In_ const wstring& lpszPropType)
{
	if (lpszPropType.empty() || PropTypeArray.empty()) return PT_UNSPECIFIED;

	// Check for numbers first before trying the string as an array lookup.
	// This will translate '0x102' to 0x102, 0x3 to 3, etc.
	auto ulType = wstringToUlong(lpszPropType, 16);
	if (ulType != NULL) return ulType;

	auto ulPropType = PT_UNSPECIFIED;

	for (auto propType : PropTypeArray)
	{
		if (0 == lstrcmpiW(lpszPropType.c_str(), propType.lpszName))
		{
			ulPropType = propType.ulValue;
			break;
		}
	}

	return ulPropType;
}

wstring GUIDToString(_In_opt_ LPCGUID lpGUID)
{
	GUID nullGUID = { 0 };

	if (!lpGUID)
	{
		lpGUID = &nullGUID;
	}

	return format(L"{%.8X-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X}", // STRING_OK
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
}

wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID)
{
	auto szGUID = GUIDToString(lpGUID);

	szGUID += L" = "; // STRING_OK

	if (lpGUID)
	{
		for (auto guid : PropGuidArray)
		{
			if (IsEqualGUID(*lpGUID, *guid.lpGuid))
			{
				return szGUID + guid.lpszName;
			}
		}
	}

	return szGUID + loadstring(IDS_UNKNOWNGUID);
}

LPCGUID GUIDNameToGUID(_In_ const wstring& szGUID, bool bByteSwapped)
{
	LPGUID lpGuidRet = nullptr;
	LPCGUID lpGUID = nullptr;
	GUID guid = { 0 };

	// Try the GUID like PS_* first
	for (auto propGuid : PropGuidArray)
	{
		if (0 == lstrcmpiW(szGUID.c_str(), propGuid.lpszName))
		{
			lpGUID = propGuid.lpGuid;
			break;
		}
	}

	if (!lpGUID) // no match - try it like a guid {}
	{
		guid = StringToGUID(szGUID, bByteSwapped);
		if (guid != GUID_NULL)
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

_Check_return_ GUID StringToGUID(_In_ const wstring& szGUID)
{
	return StringToGUID(szGUID, false);
}

_Check_return_ GUID StringToGUID(_In_ const wstring& szGUID, bool bByteSwapped)
{
	auto guid = GUID_NULL;
	if (szGUID.empty()) return guid;

	auto bin = HexStringToBin(szGUID, sizeof(GUID));
	if (bin.size() == sizeof(GUID))
	{
		memcpy(&guid, bin.data(), sizeof(GUID));

		// Note that we get the bByteSwapped behavior by default. We have to work to get the 'normal' behavior
		if (!bByteSwapped)
		{
			auto lpByte = reinterpret_cast<LPBYTE>(&guid);
			auto bByte = lpByte[0];
			lpByte[0] = lpByte[3];
			lpByte[3] = bByte;
			bByte = lpByte[1];
			lpByte[1] = lpByte[2];
			lpByte[2] = bByte;
		}
	}

	return guid;
}

// Returns string built from NameIDArray
wstring NameIDToPropName(_In_ LPMAPINAMEID lpNameID)
{
	wstring szResultString;
	if (!lpNameID) return szResultString;
	if (lpNameID->ulKind != MNID_ID) return szResultString;
	ULONG ulMatch = ulNoMatch;

	if (NameIDArray.empty()) return szResultString;

	for (ULONG ulCur = 0; ulCur < NameIDArray.size(); ulCur++)
	{
		if (NameIDArray[ulCur].lValue == lpNameID->Kind.lID)
		{
			ulMatch = ulCur;
			break;
		}
	}

	if (ulNoMatch != ulMatch)
	{
		for (auto ulCur = ulMatch; ulCur < NameIDArray.size(); ulCur++)
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

//_Check_return_ wstring InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, wstring szPrefix);

// Interprets a flag value according to a flag name and returns a string
// Will not return a string if the flag name is not recognized
wstring InterpretFlags(ULONG ulFlagName, LONG lFlagValue)
{
	ULONG ulCurEntry = 0;

	if (FlagArray.empty()) return L"";

	while (ulCurEntry < FlagArray.size() && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
	{
		ulCurEntry++;
	}

	// Don't run off the end of the array
	if (FlagArray.size() == ulCurEntry) return L"";
	if (FlagArray[ulCurEntry].ulFlagName != ulFlagName) return L"";

	// We've matched our flag name to the array - we SHOULD return a string at this point
	auto bNeedSeparator = false;

	auto lTempValue = lFlagValue;
	wstring szTempString;
	for (; FlagArray[ulCurEntry].ulFlagName == ulFlagName; ulCurEntry++)
	{
		if (flagFLAG == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue & lTempValue)
			{
				if (bNeedSeparator)
				{
					szTempString += L" | "; // STRING_OK
				}

				szTempString += FlagArray[ulCurEntry].lpszName;
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
					szTempString += L" | "; // STRING_OK
				}

				szTempString += FlagArray[ulCurEntry].lpszName;
				lTempValue = 0;
				bNeedSeparator = true;
			}
		}
		else if (flagVALUEHIGHBYTES == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 16 & 0xFFFF))
			{
				if (bNeedSeparator)
				{
					szTempString += L" | "; // STRING_OK
				}

				szTempString += FlagArray[ulCurEntry].lpszName;
				lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 16);
				bNeedSeparator = true;
			}
		}
		else if (flagVALUE3RDBYTE == FlagArray[ulCurEntry].ulFlagType)
		{
			if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 8 & 0xFF))
			{
				if (bNeedSeparator)
				{
					szTempString += L" | "; // STRING_OK
				}

				szTempString += FlagArray[ulCurEntry].lpszName;
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
					szTempString += L" | "; // STRING_OK
				}

				szTempString += FlagArray[ulCurEntry].lpszName;
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
					szTempString += L" | "; // STRING_OK
				}

				szTempString += FlagArray[ulCurEntry].lpszName;
				lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
				bNeedSeparator = true;
			}
		}
		else if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
		{
			// find any bits we need to clear
			auto lClearedBits = FlagArray[ulCurEntry].lFlagValue & lTempValue;
			// report what we found
			if (0 != lClearedBits)
			{
				if (bNeedSeparator)
				{
					szTempString += L" | "; // STRING_OK
				}

				szTempString += format(L"0x%X", lClearedBits); // STRING_OK
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
		if (bNeedSeparator)
		{
			szTempString += L" | "; // STRING_OK
		}

		szTempString += format(L"0x%X", lTempValue); // STRING_OK
	}

	return szTempString;
}

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
wstring AllFlagsToString(ULONG ulFlagName, bool bHex)
{
	wstring szFlagString;
	if (!ulFlagName) return szFlagString;
	if (FlagArray.empty()) return szFlagString;

	ULONG ulCurEntry = 0;
	wstring szTempString;

	while (ulCurEntry < FlagArray.size() && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
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
				szFlagString += formatmessage(IDS_FLAGTOSTRINGHEX, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName);
			}
			else
			{
				szFlagString += formatmessage(IDS_FLAGTOSTRINGDEC, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName);
			}
		}
	}

	return szFlagString;
}

// Returns LPSPropValue with value of a property
// Uses GetProps and falls back to OpenProperty if the value is large
// Free with MAPIFreeBuffer
_Check_return_ HRESULT GetLargeProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp)
{
	if (!lpMAPIProp || !lppProp) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric, L"GetLargeProp getting buffer from 0x%08X\n", ulPropTag);

	auto hRes = S_OK;
	ULONG cValues = 0;
	LPSPropValue lpPropArray = nullptr;
	auto bSuccess = false;

	const SizedSPropTagArray(1, sptaBuffer) =
	{
	1,
	ulPropTag
	};
	*lppProp = nullptr;

	WC_H_GETPROPS(lpMAPIProp->GetProps(LPSPropTagArray(&sptaBuffer), 0, &cValues, &lpPropArray));

	if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag) && MAPI_E_NOT_ENOUGH_MEMORY == lpPropArray->Value.err)
	{
		DebugPrint(DBGGeneric, L"GetLargeProp property reported in GetProps as large.\n");
		MAPIFreeBuffer(lpPropArray);
		lpPropArray = nullptr;
		// need to get the data as a stream
		LPSTREAM lpStream = nullptr;

		WC_MAPI(lpMAPIProp->OpenProperty(
			ulPropTag,
			&IID_IStream,
			STGM_READ,
			0,
			reinterpret_cast<LPUNKNOWN*>(&lpStream)));
		if (SUCCEEDED(hRes) && lpStream)
		{
			STATSTG StatInfo = { nullptr };
			lpStream->Stat(&StatInfo, STATFLAG_NONAME); // find out how much space we need

			// We're not going to try to support MASSIVE properties.
			if (!StatInfo.cbSize.HighPart)
			{
				EC_H(MAPIAllocateBuffer(
					sizeof(SPropValue),
					reinterpret_cast<LPVOID*>(&lpPropArray)));
				if (lpPropArray)
				{
					memset(lpPropArray, 0, sizeof(SPropValue));
					lpPropArray->ulPropTag = ulPropTag;

					if (StatInfo.cbSize.LowPart)
					{
						LPBYTE lpBuffer = nullptr;
						auto ulBufferSize = StatInfo.cbSize.LowPart;
						ULONG ulTrailingNullSize = 0;
						switch (PROP_TYPE(ulPropTag))
						{
						case PT_STRING8: ulTrailingNullSize = sizeof(char); break;
						case PT_UNICODE: ulTrailingNullSize = sizeof(WCHAR); break;
						case PT_BINARY: break;
						default: break;
						}

						EC_H(MAPIAllocateMore(
							ulBufferSize + ulTrailingNullSize,
							lpPropArray,
							reinterpret_cast<LPVOID*>(&lpBuffer)));
						if (lpBuffer)
						{
							memset(lpBuffer, 0, ulBufferSize + ulTrailingNullSize);
							ULONG ulSizeRead = 0;
							EC_MAPI(lpStream->Read(lpBuffer, ulBufferSize, &ulSizeRead));
							if (SUCCEEDED(hRes) && ulSizeRead == ulBufferSize)
							{
								switch (PROP_TYPE(ulPropTag))
								{
								case PT_STRING8:
									lpPropArray->Value.lpszA = reinterpret_cast<LPSTR>(lpBuffer);
									break;
								case PT_UNICODE:
									lpPropArray->Value.lpszW = reinterpret_cast<LPWSTR>(lpBuffer);
									break;
								case PT_BINARY:
									lpPropArray->Value.bin.cb = ulBufferSize;
									lpPropArray->Value.bin.lpb = lpBuffer;
									break;
								default: break;
								}

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
		DebugPrint(DBGGeneric, L"GetLargeProp GetProps found property.\n");
		bSuccess = true;
	}
	else if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag))
	{
		DebugPrint(DBGGeneric, L"GetLargeProp GetProps reported property as error 0x%08X.\n", lpPropArray->Value.err);
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