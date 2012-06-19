#include "stdafx.h"

#include "MrMAPI.h"
#include <shlwapi.h>
#include "guids.h"
#include "SmartView.h"
#include "InterpretProp2.h"

// Searches a NAMEID_ARRAY_ENTRY array for a target dispid.
// Exact matches are those that match
// If no hits, then ulNoMatch should be returned for lpulFirstExact
void FindNameIDArrayMatches(_In_ LONG lTarget,
							_In_count_(ulMyArray) NAMEID_ARRAY_ENTRY* MyArray,
							_In_ ULONG ulMyArray,
							_Out_ ULONG* lpulNumExacts,
							_Out_ ULONG* lpulFirstExact)
{
	ULONG ulLowerBound = 0;
	ULONG ulUpperBound = ulMyArray-1; // ulMyArray-1 is the last entry
	ULONG ulMidPoint = (ulUpperBound+ulLowerBound)/2;
	ULONG ulFirstMatch = ulNoMatch;
	ULONG ulLastMatch = ulNoMatch;

	if (lpulNumExacts) *lpulNumExacts = 0;
	if (lpulFirstExact) *lpulFirstExact = ulNoMatch;

	// find A match
	while (ulUpperBound - ulLowerBound > 1)
	{
		if (lTarget == (MyArray[ulMidPoint].lValue))
		{
			ulFirstMatch = ulMidPoint;
			break;
		}

		if (lTarget < (MyArray[ulMidPoint].lValue))
		{
			ulUpperBound = ulMidPoint;
		}
		else if (lTarget > (MyArray[ulMidPoint].lValue))
		{
			ulLowerBound = ulMidPoint;
		}
		ulMidPoint = (ulUpperBound+ulLowerBound)/2;
	}

	// When we get down to two points, we may have only checked one of them
	// Make sure we've checked the other
	if (lTarget == (MyArray[ulUpperBound].lValue))
	{
		ulFirstMatch = ulUpperBound;
	}
	else if (lTarget == (MyArray[ulLowerBound].lValue))
	{
		ulFirstMatch = ulLowerBound;
	}

	// Check that we got a match
	if (ulNoMatch != ulFirstMatch)
	{
		ulLastMatch = ulFirstMatch; // Remember the last match we've found so far

		// Scan backwards to find the first match
		while (ulFirstMatch > 0 && lTarget == MyArray[ulFirstMatch-1].lValue)
		{
			ulFirstMatch = ulFirstMatch - 1;
		}

		// Scan forwards to find the real last match
		// Last entry in the array is ulPropTagArray-1
		while (ulLastMatch+1 < ulMyArray && lTarget == MyArray[ulLastMatch+1].lValue)
		{
			ulLastMatch = ulLastMatch + 1;
		}

		ULONG ulNumMatches = 0;

		if (ulNoMatch != ulFirstMatch)
		{
			ulNumMatches = ulLastMatch - ulFirstMatch + 1;
		}

		if (lpulNumExacts) *lpulNumExacts = ulNumMatches;
		if (lpulFirstExact) *lpulFirstExact = ulFirstMatch;
	}
} // FindNameIDArrayMatches

// prints the type of a prop tag
// no pretty stuff or \n - calling function gets to do that
void PrintType(_In_ ULONG ulPropTag)
{
	bool bNeedInstance = false;

	if (ulPropTag & MV_INSTANCE)
	{
		ulPropTag &= ~MV_INSTANCE;
		bNeedInstance = true;
	}

	// Do a linear search through PropTypeArray - ulPropTypeArray will be small
	ULONG ulMatch = 0;
	bool bFound = false;

	for (ulMatch = 0 ; ulMatch < ulPropTypeArray ; ulMatch++)
	{
		if (PROP_TYPE(ulPropTag) == PropTypeArray[ulMatch].ulValue)
		{
			bFound = true;
			printf("%ws",PropTypeArray[ulMatch].lpszName);
			break;
		}
	}

	if (!bFound) printf("0x%04X = Unknown type",PROP_TYPE(ulPropTag));

	if (bNeedInstance) printf(" | MV_INSTANCE");
} // PrintType

void PrintKnownTypes()
{
	// Do a linear search through PropTypeArray - ulPropTypeArray will be small
	ULONG ulMatch = 0;

	printf("%-18s%-9s%s\n","Type","Hex","Decimal");
	for (ulMatch = 0 ; ulMatch < ulPropTypeArray ; ulMatch++)
	{
		printf("%-15ws = 0x%04X = %d\n",PropTypeArray[ulMatch].lpszName,PropTypeArray[ulMatch].ulValue,PropTypeArray[ulMatch].ulValue);
	}

	printf("\n");
	printf("Types may also have the flag 0x%04X = %s\n",
		MV_INSTANCE,"MV_INSTANCE");
} // PrintKnownTypes

// Print the tag found in the array at ulRow
void PrintTag(_In_ ULONG ulRow)
{
	printf("0x%08X,%ws,",PropTagArray[ulRow].ulValue,PropTagArray[ulRow].lpszName);
	PrintType(PropTagArray[ulRow].ulValue);
	printf("\n");
} // PrintTag

// Given a property tag, output the matching names and partially matching names
// Matching names will be presented first, with partial matches following
// Example:
// C:\>MrMAPI 0x3001001F
// Property tag 0x3001001F:
//
// Exact matches:
// 0x3001001F,PR_DISPLAY_NAME_W,PT_UNICODE
// 0x3001001F,PidTagDisplayName,PT_UNICODE
//
// Partial matches:
// 0x3001001E,PR_DISPLAY_NAME,PT_STRING8
// 0x3001001E,PR_DISPLAY_NAME_A,PT_STRING8
// 0x3001001E,ptagDisplayName,PT_STRING8
void PrintTagFromNum(_In_ ULONG ulPropTag)
{
	ULONG ulPropID = NULL;
	ULONG ulNumExacts = NULL;
	ULONG ulFirstExactMatch = ulNoMatch;
	ULONG ulNumPartials = NULL;
	ULONG ulFirstPartial = ulNoMatch;

	if (ulPropTag & 0xffff0000) // dealing with a full prop tag
	{
		ulPropID = PROP_ID(ulPropTag);
	}
	else
	{
		ulPropID = ulPropTag;
	}

	printf("Property tag 0x%08X:\n",ulPropTag);

	if (ulPropID & 0x8000)
	{
		printf("Since this property ID is greater than 0x8000 there is a good chance it is a named property.\n");
		printf("Only trust the following output if this property is known to be an address book property.\n");
	}

	FindTagArrayMatches(ulPropTag,true,PropTagArray,ulPropTagArray,&ulNumExacts,&ulFirstExactMatch,&ulNumPartials,&ulFirstPartial);

	ULONG ulCur = NULL;
	if (ulNumExacts > 0)
	{
		printf("\nExact matches:\n");
		for (ulCur = ulFirstExactMatch ; ulCur < ulFirstExactMatch+ulNumExacts ; ulCur++)
		{
			PrintTag(ulCur);
		}
	}

	if (ulNumPartials > 0)
	{
		printf("\nPartial matches:\n");
		// let's print our partial matches
		for (ulCur = ulFirstPartial ; ulCur < ulFirstPartial+ulNumPartials+ulNumExacts ; ulCur++)
		{
			if (ulPropTag == PropTagArray[ulCur].ulValue) continue; // skip our exact matches
			PrintTag(ulCur);
		}
	}
} // PrintTagFromNum

void PrintTagFromName(_In_z_ LPCWSTR lpszPropName)
{
	if (!lpszPropName) return;

	ULONG ulCur = 0;
	bool bMatchFound = false;

	for (ulCur = 0 ; ulCur < ulPropTagArray ; ulCur++)
	{
		if (0 == lstrcmpiW(lpszPropName,PropTagArray[ulCur].lpszName))
		{
			PrintTag(ulCur);

			// now that we have a match, let's see if we have other tags with the same number
			ULONG ulExactMatch = ulCur; // The guy that matched lpszPropName
			ULONG ulNumExacts = NULL;
			ULONG ulFirstExactMatch = ulNoMatch;
			ULONG ulNumPartials = NULL;
			ULONG ulFirstPartial = ulNoMatch;

			FindTagArrayMatches(PropTagArray[ulExactMatch].ulValue,true,PropTagArray,ulPropTagArray,&ulNumExacts,&ulFirstExactMatch,&ulNumPartials,&ulFirstPartial);

			// We're gonna skip at least one, so only print if we have more than one
			if (ulNumExacts > 1)
			{
				printf("\nOther exact matches:\n");
				for (ulCur = ulFirstExactMatch ; ulCur < ulFirstExactMatch+ulNumExacts ; ulCur++)
				{
					if (ulExactMatch == ulCur) continue; // skip this one
					PrintTag(ulCur);
				}
			}

			if (ulNumPartials > 0)
			{
				printf("\nOther partial matches:\n");
				for (ulCur = ulFirstPartial ; ulCur < ulFirstPartial+ulNumPartials+ulNumExacts ; ulCur++)
				{
					if (PropTagArray[ulExactMatch].ulValue == PropTagArray[ulCur].ulValue) continue; // skip our exact matches
					PrintTag(ulCur);
				}
			}
			bMatchFound = true;
			break;
		}
	}

	if (!bMatchFound) printf("Property tag \"%ws\" not found\n",lpszPropName);
} // PrintTagFromName

// Search for properties matching lpszPropName on a substring
// If ulType isn't ulNoMatch, restrict on the property type as well
void PrintTagFromPartialName(_In_opt_z_ LPCWSTR lpszPropName, _In_ ULONG ulType)
{
	if (lpszPropName) printf("Searching for \"%ws\"\n",lpszPropName);
	else printf("Searching for all properties\n");

	if (ulNoMatch != ulType)
	{
		printf("Restricting output to ");
		PrintType(ulType);
		printf("\n");
	}
	ULONG ulCur = 0;
	ULONG ulNumMatches = 0;

	for (ulCur = 0 ; ulCur < ulPropTagArray ; ulCur++)
	{
		if (!lpszPropName || 0 != StrStrIW(PropTagArray[ulCur].lpszName,lpszPropName))
		{
			if (ulNoMatch != ulType && ulType != PROP_TYPE(PropTagArray[ulCur].ulValue)) continue;
			PrintTag(ulCur);
			ulNumMatches++;
		}
	}
	printf("Found %d matches.\n",ulNumMatches);
} // PrintTagFromPartialName

void PrintGUID(_In_ LPCGUID lpGUID)
{
	// PSUNKNOWN isn't a real guid - just a placeholder - don't print it
	if (!lpGUID || IsEqualGUID(*lpGUID,PSUNKNOWN))
	{
		printf(",");
		return;
	}

	printf("{%.8X-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X}",
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

	ULONG	ulCur = 0;

	printf(",");
	if (ulPropGuidArray && PropGuidArray)
	{
		for (ulCur = 0 ; ulCur < ulPropGuidArray ; ulCur++)
		{
			if (IsEqualGUID(*lpGUID,*PropGuidArray[ulCur].lpGuid))
			{
				printf("%ws", PropGuidArray[ulCur].lpszName);
				break;
			}
		}
	}
} // PrintGUID

void PrintGUIDs()
{
	ULONG	ulCur = 0;
	for (ulCur = 0 ; ulCur < ulPropGuidArray ; ulCur++)
	{
		printf("{%.8X-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X}",
			PropGuidArray[ulCur].lpGuid->Data1,
			PropGuidArray[ulCur].lpGuid->Data2,
			PropGuidArray[ulCur].lpGuid->Data3,
			PropGuidArray[ulCur].lpGuid->Data4[0],
			PropGuidArray[ulCur].lpGuid->Data4[1],
			PropGuidArray[ulCur].lpGuid->Data4[2],
			PropGuidArray[ulCur].lpGuid->Data4[3],
			PropGuidArray[ulCur].lpGuid->Data4[4],
			PropGuidArray[ulCur].lpGuid->Data4[5],
			PropGuidArray[ulCur].lpGuid->Data4[6],
			PropGuidArray[ulCur].lpGuid->Data4[7]);
		printf(",%ws\n", PropGuidArray[ulCur].lpszName);
	}
} // PrintGUIDs

void PrintDispID(_In_ ULONG ulRow)
{
	printf("0x%04X,%ws,",NameIDArray[ulRow].lValue,NameIDArray[ulRow].lpszName);
	PrintGUID(NameIDArray[ulRow].lpGuid);
	printf(",");
	if (PT_UNSPECIFIED != NameIDArray[ulRow].ulType)
	{
		PrintType(PROP_TAG(NameIDArray[ulRow].ulType,0));
	}
	printf(",");
	if (NameIDArray[ulRow].lpszArea)
	{
		printf("%ws",NameIDArray[ulRow].lpszArea);
	}
	printf("\n");
} // PrintDispID

void PrintDispIDFromNum(_In_ ULONG ulDispID)
{
	ULONG ulNumExacts = NULL;
	ULONG ulFirstExactMatch = ulNoMatch;

	printf("Dispid tag 0x%04X:\n",ulDispID);

	FindNameIDArrayMatches(ulDispID,NameIDArray,ulNameIDArray,&ulNumExacts,&ulFirstExactMatch);

	ULONG ulCur = NULL;
	if (ulNumExacts > 0 && ulNoMatch != ulFirstExactMatch)
	{
		printf("\nExact matches:\n");
		for (ulCur = ulFirstExactMatch ; ulCur < ulFirstExactMatch+ulNumExacts ; ulCur++)
		{
			PrintDispID(ulCur);
		}
	}
} // PrintDispIDFromNum

void PrintDispIDFromName(_In_z_ LPCWSTR lpszDispIDName)
{
	if (!lpszDispIDName) return;

	ULONG ulCur = 0;
	bool bMatchFound = false;

	for (ulCur = 0 ; ulCur < ulNameIDArray ; ulCur++)
	{
		if (0 == lstrcmpiW(lpszDispIDName,NameIDArray[ulCur].lpszName))
		{
			PrintDispID(ulCur);

			// now that we have a match, let's see if we have other dispids with the same number
			ULONG ulExactMatch = ulCur; // The guy that matched lpszPropName
			ULONG ulNumExacts = NULL;
			ULONG ulFirstExactMatch = ulNoMatch;

			FindNameIDArrayMatches(NameIDArray[ulExactMatch].lValue,NameIDArray,ulNameIDArray,&ulNumExacts,&ulFirstExactMatch);

			// We're gonna skip at least one, so only print if we have more than one
			if (ulNumExacts > 1)
			{
				printf("\nOther exact matches:\n");
				for (ulCur = ulFirstExactMatch ; ulCur < ulFirstExactMatch+ulNumExacts ; ulCur++)
				{
					if (ulExactMatch == ulCur) continue; // skip this one
					PrintDispID(ulCur);
				}
			}
			bMatchFound = true;
			break;
		}
	}

	if (!bMatchFound) printf("Property tag \"%ws\" not found\n",lpszDispIDName);
} // PrintDispIDFromName

// Search for properties matching lpszPropName on a substring
void PrintDispIDFromPartialName(_In_opt_z_ LPCWSTR lpszDispIDName, _In_ ULONG ulType)
{
	if (lpszDispIDName) printf("Searching for \"%ws\"\n",lpszDispIDName);
	else printf("Searching for all properties\n");

	ULONG ulCur = 0;
	ULONG ulNumMatches = 0;

	for (ulCur = 0 ; ulCur < ulNameIDArray ; ulCur++)
	{
		if (!lpszDispIDName || 0 != StrStrIW(NameIDArray[ulCur].lpszName,lpszDispIDName))
		{
			if (ulNoMatch != ulType && ulType != NameIDArray[ulCur].ulType) continue;
			PrintDispID(ulCur);
			ulNumMatches++;
		}
	}
	printf("Found %d matches.\n",ulNumMatches);
} // PrintDispIDFromPartialName

void PrintFlag(_In_ ULONG ulPropNum, _In_opt_z_ LPCWSTR lpszPropName, _In_ bool bIsDispid, _In_ ULONG ulFlagValue)
{
	LPWSTR szFlags = NULL;
	ULONG ulCur = 0;
	if (bIsDispid)
	{
		if (ulPropNum)
		{
			for (ulCur = 0 ; ulCur < ulNameIDArray ; ulCur++)
			{
				if (ulPropNum == (ULONG) NameIDArray[ulCur].lValue)
				{
					break;
				}
			}
		}
		else if (lpszPropName)
		{
			for (ulCur = 0 ; ulCur < ulNameIDArray ; ulCur++)
			{
				if (NameIDArray[ulCur].lpszName && 0 == lstrcmpiW(lpszPropName,NameIDArray[ulCur].lpszName))
				{
					break;
				}
			}
		}
		if (ulCur != ulNameIDArray)
		{
			printf("Found named property %ws (0x%04X) ",NameIDArray[ulCur].lpszName, NameIDArray[ulCur].lValue);
			PrintGUID(NameIDArray[ulCur].lpGuid);
			printf(".\n");
			InterpretNumberAsStringNamedProp(ulFlagValue,NameIDArray[ulCur].lValue, NameIDArray[ulCur].lpGuid,&szFlags);
		}
	}
	else
	{
		if (ulPropNum)
		{
			for (ulCur = 0 ; ulCur < ulPropTagArray ; ulCur++)
			{
				if (ulPropNum == PropTagArray[ulCur].ulValue)
				{
					break;
				}
			}
		}
		else if (lpszPropName)
		{
			for (ulCur = 0 ; ulCur < ulPropTagArray ; ulCur++)
			{
				if (0 == lstrcmpiW(lpszPropName,PropTagArray[ulCur].lpszName))
				{
					break;
				}
			}
		}


		if (ulCur != ulPropTagArray)
		{
			printf("Found property %ws (0x%08X).\n",PropTagArray[ulCur].lpszName, PropTagArray[ulCur].ulValue);
			InterpretNumberAsStringProp(ulFlagValue,PropTagArray[ulCur].ulValue,&szFlags);
		}
	}
	if (szFlags)
	{
		printf("0x%08X = %ws\n",ulFlagValue,szFlags);
		delete[] szFlags;
	}
	else
	{
		printf("No flag parsing found.\n");
	}
} // PrintFlag

void DoPropTags(_In_ MYOPTIONS ProgOpts)
{
	ULONG ulPropNum = NULL;
	LPWSTR lpszPropName = ProgOpts.lpszUnswitchedOption;

	if (lpszPropName)
	{
		ULONG ulArg = NULL;
		LPWSTR szEndPtr = NULL;
		ulArg = wcstoul(lpszPropName,&szEndPtr, (ProgOpts.ulOptions & OPT_DODECIMAL)?10:16);

		// if szEndPtr is pointing to something other than NULL, this must be a string
		if (!szEndPtr || *szEndPtr)
		{
			ulArg = NULL;
		}

		ulPropNum = ulArg;
	}

	// Handle dispid cases
	if (ProgOpts.ulOptions & OPT_DODISPID)
	{
		if (ProgOpts.ulOptions & OPT_DOFLAG)
		{
			PrintFlag(ulPropNum, lpszPropName, true, ProgOpts.ulFlagValue);
		}
		else if (ProgOpts.ulOptions & OPT_DOPARTIALSEARCH)
		{
			PrintDispIDFromPartialName(lpszPropName,ProgOpts.ulTypeNum);
		}
		else if (ulPropNum)
		{
			PrintDispIDFromNum(ulPropNum);
		}
		else if (lpszPropName)
		{
			PrintDispIDFromName(lpszPropName);
		}
		return;
	}

	// Handle prop tag cases
	if (ProgOpts.ulOptions & OPT_DOFLAG)
	{
		PrintFlag(ulPropNum, lpszPropName, false, ProgOpts.ulFlagValue);
	}
	else if (ProgOpts.ulOptions & OPT_DOPARTIALSEARCH)
	{
		PrintTagFromPartialName(lpszPropName,ProgOpts.ulTypeNum);
	}
	// If we weren't asked about a property, maybe we were asked about types
	else if (ProgOpts.ulOptions & OPT_DOTYPE)
	{
		if (ulNoMatch != ProgOpts.ulTypeNum)
		{
			PrintType(PROP_TAG(ProgOpts.ulTypeNum,0));
			printf(" = 0x%04X = %d",ProgOpts.ulTypeNum,ProgOpts.ulTypeNum);
			printf("\n");
		}
		else
		{
			PrintKnownTypes();
		}
	}
	else
	{
		if (ulPropNum)
		{
			PrintTagFromNum(ulPropNum);
		}
		else if (lpszPropName)
		{
			PrintTagFromName(lpszPropName);
		}
	}
} // DoPropTags

void DoGUIDs(_In_ MYOPTIONS /*ProgOpts*/)
{
	PrintGUIDs();
} // DoGUIDs

void DoFlagSearch(_In_ MYOPTIONS ProgOpts)
{
	//ProgOpts.lpszFlagName;
	ULONG	ulCurEntry = 0;
	if (!ulFlagArray || !FlagArray) return;

	for (ulCurEntry = 0; ulCurEntry < ulFlagArray; ulCurEntry++)
	{
		if (!_wcsicmp(FlagArray[ulCurEntry].lpszName, ProgOpts.lpszFlagName))
		{
			printf("%ws = 0x%08X\n",FlagArray[ulCurEntry].lpszName, FlagArray[ulCurEntry].lFlagValue);
			break;
		}
	}
} // DoFlagSearch