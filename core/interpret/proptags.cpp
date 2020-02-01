#include <core/stdafx.h>
#include <core/interpret/proptags.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/utility/strings.h>
#include <core/interpret/proptype.h>
#include <core/utility/output.h>
#include <core/utility/registry.h>

namespace proptags
{
	static WCHAR szPropSeparator[] = L", "; // STRING_OK
	std::unordered_map<ULONG64, PropTagNames> g_PropNames;

	std::wstring TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine)
	{
		std::wstring szTemp;

		auto namePropNames = cache::NameIDToStrings(ulPropTag, lpObj, nullptr, nullptr, bIsAB);

		auto propTagNames = PropTagToPropName(ulPropTag, bIsAB);

		std::wstring szFormatString;
		if (bSingleLine)
		{
			szFormatString = L"0x%1!08X! (%2)"; // STRING_OK
			if (!propTagNames.bestGuess.empty()) szFormatString += L": %3!ws!"; // STRING_OK
			if (!propTagNames.otherMatches.empty()) szFormatString += L": (%4!ws!)"; // STRING_OK
			if (!namePropNames.name.empty())
			{
				szFormatString += strings::loadstring(IDS_NAMEDPROPSINGLELINE);
			}

			if (!namePropNames.guid.empty())
			{
				szFormatString += strings::loadstring(IDS_GUIDSINGLELINE);
			}
		}
		else
		{
			szFormatString = strings::loadstring(IDS_TAGMULTILINE);
			if (!propTagNames.bestGuess.empty())
			{
				szFormatString += strings::loadstring(IDS_PROPNAMEMULTILINE);
			}

			if (!propTagNames.otherMatches.empty())
			{
				szFormatString += strings::loadstring(IDS_OTHERNAMESMULTILINE);
			}

			if (PROP_ID(ulPropTag) < 0x8000)
			{
				szFormatString += strings::loadstring(IDS_DASLPROPTAG);
			}
			else if (!namePropNames.dasl.empty())
			{
				szFormatString += strings::loadstring(IDS_DASLNAMED);
			}

			if (!namePropNames.name.empty())
			{
				szFormatString += strings::loadstring(IDS_NAMEPROPNAMEMULTILINE);
			}

			if (!namePropNames.guid.empty())
			{
				szFormatString += strings::loadstring(IDS_NAMEPROPGUIDMULTILINE);
			}
		}

		auto szRet = strings::formatmessage(
			szFormatString.c_str(),
			ulPropTag,
			proptype::TypeToString(ulPropTag).c_str(),
			propTagNames.bestGuess.c_str(),
			propTagNames.otherMatches.c_str(),
			namePropNames.name.c_str(),
			namePropNames.guid.c_str(),
			namePropNames.dasl.c_str());

		if (fIsSet(output::DBGTest))
		{
			static size_t cchMaxBuff = 0;
			const auto cchBuff = szRet.length();
			cchMaxBuff = max(cchBuff, cchMaxBuff);
			output::DebugPrint(
				output::DBGTest,
				L"TagToString parsing 0x%08X returned %u chars - max %u\n",
				ulPropTag,
				static_cast<UINT>(cchBuff),
				static_cast<UINT>(cchMaxBuff));
		}

		return szRet;
	}

	PropTagNames PropTagToPropName(ULONG ulPropTag, bool bIsAB)
	{
		auto ulKey = (bIsAB ? static_cast<ULONG64>(1) << 32 : 0) | ulPropTag;

		const auto match = g_PropNames.find(ulKey);
		if (match != g_PropNames.end())
		{
			return match->second;
		}

		std::vector<ULONG> ulExacts;
		std::vector<ULONG> ulPartials;
		FindTagArrayMatches(ulPropTag, bIsAB, PropTagArray, ulExacts, ulPartials);

		PropTagNames entry;

		if (!ulExacts.empty())
		{
			entry.bestGuess = PropTagArray[ulExacts.front()].lpszName;
			ulExacts.erase(ulExacts.begin());

			for (const auto& ulMatch : ulExacts)
			{
				if (!entry.otherMatches.empty())
				{
					entry.otherMatches += szPropSeparator;
				}

				entry.otherMatches += PropTagArray[ulMatch].lpszName;
			}
		}

		if (!ulPartials.empty())
		{
			if (entry.bestGuess.empty())
			{
				entry.bestGuess = PropTagArray[ulPartials.front()].lpszName;
				ulPartials.erase(ulPartials.begin());
			}

			for (const auto& ulMatch : ulPartials)
			{
				if (!entry.otherMatches.empty())
				{
					entry.otherMatches += szPropSeparator;
				}

				entry.otherMatches += PropTagArray[ulMatch].lpszName;
			}
		}

		g_PropNames.insert({ulKey, entry});

		return entry;
	}

	// Compare tag sort order.
	bool CompareTagsSortOrder(int a1, int a2) noexcept
	{
		const auto lpTag1 = &PropTagArray[a1];
		const auto lpTag2 = &PropTagArray[a2];

		if (lpTag1->ulSortOrder < lpTag2->ulSortOrder) return false;
		if (lpTag1->ulSortOrder == lpTag2->ulSortOrder)
		{
			return wcscmp(lpTag1->lpszName, lpTag2->lpszName) <= 0;
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
	void FindTagArrayMatches(
		_In_ ULONG ulTarget,
		bool bIsAB,
		const std::vector<NAME_ARRAY_ENTRY_V2>& MyArray,
		std::vector<ULONG>& ulExacts,
		std::vector<ULONG>& ulPartials)
	{
		if (!(ulTarget & PROP_TAG_MASK)) // not dealing with a full prop tag
		{
			ulTarget = PROP_TAG(PT_UNSPECIFIED, ulTarget);
		}

		ULONG ulLowerBound = 0;
		auto ulUpperBound = static_cast<ULONG>(MyArray.size() - 1); // size-1 is the last entry
		auto ulMidPoint = (ulUpperBound + ulLowerBound) / 2;
		ULONG ulFirstMatch = cache::ulNoMatch;
		const auto ulMaskedTarget = ulTarget & PROP_TAG_MASK;

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
		if (cache::ulNoMatch != ulFirstMatch)
		{
			// Scan backwards to find the first partial match
			while (ulFirstMatch > 0 && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulFirstMatch - 1].ulValue))
			{
				ulFirstMatch = ulFirstMatch - 1;
			}

			for (auto ulCur = ulFirstMatch;
				 ulCur < MyArray.size() && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulCur].ulValue);
				 ulCur++)
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

			if (!ulExacts.empty()) sort(ulExacts.begin(), ulExacts.end(), CompareTagsSortOrder);
			if (!ulPartials.empty()) sort(ulPartials.begin(), ulPartials.end(), CompareTagsSortOrder);
		}
	}

	// Strictly does a lookup in the array. Does not convert otherwise
	_Check_return_ ULONG LookupPropName(_In_ const std::wstring& lpszPropName)
	{
		auto trimName = strings::trim(lpszPropName);
		if (trimName.empty()) return 0;

		for (auto& tag : PropTagArray)
		{
			if (0 == lstrcmpiW(trimName.c_str(), tag.lpszName))
			{
				return tag.ulValue;
			}
		}

		return 0;
	}

	_Check_return_ ULONG PropNameToPropTag(_In_ const std::wstring& lpszPropName)
	{
		if (lpszPropName.empty()) return 0;

		const auto ulTag = strings::wstringToUlong(lpszPropName, 16);
		if (ulTag != NULL)
		{
			return ulTag;
		}

		return LookupPropName(lpszPropName);
	}
} // namespace proptags