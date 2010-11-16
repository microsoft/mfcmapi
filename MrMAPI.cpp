#include "stdafx.h"

#include "MrMAPI.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "PropTagArray.h"
#include <shlwapi.h>
#include <io.h>
#include "guids.h"
#include "SmartView.h"

// Searches an array for a target number.
// Search is done with a mask
// Partial matches are those that match with the mask applied
// Exact matches are those that match without the mask applied
// lpUlNumPartials will exclude count of exact matches
// if it want just the true partial matches.
// If no hits, then ulNoMatch should be returned for lpulFirstExact and/or lpulFirstPartial
void FindTagArrayMatches(ULONG ulTarget,
						 NAME_ARRAY_ENTRY* MyArray,
						 ULONG ulMyArray,
						 ULONG* lpulNumExacts,
						 ULONG* lpulFirstExact,
						 ULONG* lpulNumPartials,
						 ULONG* lpulFirstPartial)
{
	ULONG ulLowerBound = 0;
	ULONG ulUpperBound = ulMyArray-1; // ulMyArray-1 is the last entry
	ULONG ulMidPoint = (ulUpperBound+ulLowerBound)/2;
	ULONG ulFirstMatch = ulNoMatch;
	ULONG ulLastMatch = ulNoMatch;
	ULONG ulFirstExactMatch = ulNoMatch;
	ULONG ulLastExactMatch = ulNoMatch;
	ULONG ulMaskedTarget = ulTarget & 0xffff0000;

	if (lpulNumExacts) *lpulNumExacts = 0;
	if (lpulFirstExact) *lpulFirstExact = ulNoMatch;
	if (lpulNumPartials) *lpulNumPartials = 0;
	if (lpulFirstPartial) *lpulFirstPartial = ulNoMatch;

	// find A partial match
	while (ulUpperBound - ulLowerBound > 1)
	{
		if (ulMaskedTarget == (0xffff0000 & MyArray[ulMidPoint].ulValue))
		{
			ulFirstMatch = ulMidPoint;
			break;
		}

		if (ulMaskedTarget < (0xffff0000 & MyArray[ulMidPoint].ulValue))
		{
			ulUpperBound = ulMidPoint;
		}
		else if (ulMaskedTarget > (0xffff0000 & MyArray[ulMidPoint].ulValue))
		{
			ulLowerBound = ulMidPoint;
		}
		ulMidPoint = (ulUpperBound+ulLowerBound)/2;
	}

	// When we get down to two points, we may have only checked one of them
	// Make sure we've checked the other
	if (ulMaskedTarget == (0xffff0000 & MyArray[ulUpperBound].ulValue))
	{
		ulFirstMatch = ulUpperBound;
	}
	else if (ulMaskedTarget == (0xffff0000 & MyArray[ulLowerBound].ulValue))
	{
		ulFirstMatch = ulLowerBound;
	}

	// Check that we got a match
	if (ulNoMatch != ulFirstMatch)
	{
		ulLastMatch = ulFirstMatch; // Remember the last match we've found so far

		// Scan backwards to find the first partial match
		while (ulFirstMatch > 0 && ulMaskedTarget == (0xffff0000 & MyArray[ulFirstMatch-1].ulValue))
		{
			ulFirstMatch = ulFirstMatch - 1;
		}

		// Scan forwards to find the real last partial match
		// Last entry in the array is ulPropTagArray-1
		while (ulLastMatch+1 < ulPropTagArray && ulMaskedTarget == (0xffff0000 & MyArray[ulLastMatch+1].ulValue))
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
						 
// Searches a NAMEID_ARRAY_ENTRY array for a target dispid.
// Exact matches are those that match
// If no hits, then ulNoMatch should be returned for lpulFirstExact
void FindNameIDArrayMatches(LONG lTarget,
							NAMEID_ARRAY_ENTRY* MyArray,
							ULONG ulMyArray,
							ULONG* lpulNumExacts,
							ULONG* lpulFirstExact)
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
		while (ulLastMatch+1 < ulPropTagArray && lTarget == MyArray[ulLastMatch+1].lValue)
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
void PrintType(ULONG ulPropTag)
{
	bool bNeedInstance = false;

	if (ulPropTag & MV_INSTANCE)
	{
		ulPropTag &= ~MV_INSTANCE;
		bNeedInstance = true;
	}

	// Do a linear search through PropTypeArray - ulPropTypeArray will be small
	ULONG ulMatch = 0;
	BOOL bFound = false;

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

	if (bNeedInstance) printf(_T(" | MV_INSTANCE"));
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
void PrintTag(ULONG ulRow)
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
void PrintTagFromNum(ULONG ulPropTag)
{
	ULONG ulPropID = NULL;
	ULONG ulPropType = NULL;
	ULONG ulNumExacts = NULL;
	ULONG ulFirstExactMatch = ulNoMatch;
	ULONG ulNumPartials = NULL;
	ULONG ulFirstPartial = ulNoMatch;

	// determine the prop ID we're seeking
	if (ulPropTag & 0xffff0000) // dealing with a full prop tag
	{
		ulPropID   = PROP_ID(ulPropTag);
		ulPropType = PROP_TYPE(ulPropTag);
	}
	else
	{
		ulPropID   = ulPropTag;
		ulPropType = PT_UNSPECIFIED;
	}

	// put everything back together
	ulPropTag = PROP_TAG(ulPropType,ulPropID);

	printf("Property tag 0x%08X:\n",ulPropTag);

	if (ulPropID & 0x8000)
	{
		printf("Since this property ID is greater than 0x8000 there is a good chance it is a named property.\n");
		printf("Only trust the following output if this property is known to be an address book property.\n");
	}

	FindTagArrayMatches(ulPropTag,PropTagArray,ulPropTagArray,&ulNumExacts,&ulFirstExactMatch,&ulNumPartials,&ulFirstPartial);

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

void PrintTagFromName(LPCWSTR lpszPropName)
{
	if (!lpszPropName) return;

	ULONG ulCur = 0;
	BOOL bMatchFound = false;

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

			FindTagArrayMatches(PropTagArray[ulExactMatch].ulValue,PropTagArray,ulPropTagArray,&ulNumExacts,&ulFirstExactMatch,&ulNumPartials,&ulFirstPartial);

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
void PrintTagFromPartialName(LPCWSTR lpszPropName, ULONG ulType)
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

void PrintGUID(LPCGUID lpGUID)
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

void PrintDispID(ULONG ulRow)
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

void PrintDispIDFromNum(ULONG ulDispID)
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
} // PrintTagFromNum

void PrintDispIDFromName(LPCWSTR lpszDispIDName)
{
	if (!lpszDispIDName) return;

	ULONG ulCur = 0;
	BOOL bMatchFound = false;

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
void PrintDispIDFromPartialName(LPCWSTR lpszDispIDName, ULONG ulType)
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

void DisplayUsage()
{
	printf("Usage:\n");
	printf("   MrMAPI [-S] [-D] [-N] [-T <type>] <number>|<name>\n");
	printf("   MrMAPI -P <type> -I <input file> [-O <output file>]\n");
	printf("\n");
	printf("   Property Tag Lookup:\n");
	printf("   -S   Perform substring search.\n");
	printf("           With no parameters prints all known properties.\n");
	printf("   -D   Search dispids.\n");
	printf("   -N   Number is in decimal. Ignored for non-numbers.\n");
	printf("   -T   Print information on specified type.\n");
	printf("           With no parameters prints list of known types.\n");
	printf("           When combined with -S, restrict output to given type.\n");
	printf("\n");
	printf("   Smart View Parsing:\n");
	printf("   -P   Parser type (number). See list below for supported parsers.\n");
	printf("   -I   Input file.\n");
	printf("   -O   Output file (optional).\n");
	printf("\n");
	printf("Smart View Parsers:\n");
	// Print smart view options
	ULONG i = 1;
	for (i = 1 ; i < g_cuidParsingTypes ; i++)
	{
		HRESULT hRes = S_OK;
		CString szStruct;
		EC_B(szStruct.LoadString(g_uidParsingTypes[i]));
		printf("   %2d %s\n",i,(LPCTSTR) szStruct);
	}
	printf("\n");
	printf("Examples:\n");
	printf("   MrMAPI PR_DISPLAY_NAME\n");
	printf("\n");
	printf("   MrMAPI 0x3001001e\n");
	printf("   MrMAPI 3001001e\n");
	printf("   MrMAPI 3001\n");
	printf("\n");
	printf("   MrMAPI -n 12289\n");
	printf("\n");
	printf("   MrMAPI -t PT_LONG\n");
	printf("   MrMAPI -t 3102\n");
	printf("   MrMAPI -t\n");
	printf("\n");
	printf("   MrMAPI -s display\n");
	printf("   MrMAPI -s display -t PT_LONG\n");
	printf("   MrMAPI -t 102 -s display\n");
	printf("\n");
	printf("   MrMAPI -d dispidReminderTime\n");
	printf("   MrMAPI -d 0x8502\n");
	printf("   MrMAPI -d -s reminder\n");
	printf("   MrMAPI -d -n 34050\n");
	printf("\n");
	printf("   MrMAPI -p 17 -i webview.txt -o parsed.txt\n");
	printf("\n");
	printf("   There are currently %d known property tags, %d known dispids, %d known types, and %d smart view parsers.\n",ulPropTagArray,ulNameIDArray,ulPropTypeArray,g_cuidParsingTypes-1);
} // DisplayUsage

// scans an arg and returns the string or hex number that it represents
void GetArg(char* szArgIn, WCHAR** lpszArgOut, ULONG* ulArgOut)
{
	ULONG ulArg = NULL;
	WCHAR* szArg = NULL;
	LPSTR szEndPtr = NULL;
	ulArg = strtoul(szArgIn,&szEndPtr,16);

	// if szEndPtr is pointing to something other than NULL, this must be a string
	if (!szEndPtr || *szEndPtr)
	{
		HRESULT hRes = S_OK;
		ulArg = NULL;
		EC_H(AnsiToUnicode(szArgIn,&szArg));
	}

	if (lpszArgOut) *lpszArgOut = szArg;
	if (ulArgOut)   *ulArgOut   = ulArg;
} // GetArg

BOOL bSetMode(CmdMode* pMode, CmdMode TargetMode)
{
	if (pMode && ((cmdmodeUnknown == *pMode) || (TargetMode == *pMode)))
	{
		*pMode = TargetMode;
		return true;
	}
	return false;
} // bSetMode

// Parses command line arguments and fills out MYOPTIONS
BOOL ParseArgs(int argc, char * argv[], MYOPTIONS * pRunOpts)
{
	HRESULT hRes = S_OK;
	// Clear our options list
	ZeroMemory(pRunOpts,sizeof(MYOPTIONS));

	pRunOpts->ulTypeNum = ulNoMatch;

	if (!pRunOpts) return false;
	if (1 == argc) return false;

	for (int i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '-':
		case '/':
		case '\\':
			{
				if (0 == argv[i][1])
				{
					// Bad argument - get out of here
					return false;
				}
				switch (tolower(argv[i][1]))
				{
				// Proptag parsing
				case 's':
					if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
					pRunOpts->bDoPartialSearch = true;
					break;
				case 'd':
					if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
					pRunOpts->bDoDispid = true;
					break;
				case 'n':
					if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
					pRunOpts->bDoDecimal = true;
					break;
				case 't':
					if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
					pRunOpts->bDoType = true;
					if (i+1 < argc)
					{
						GetArg(argv[i+1],&pRunOpts->lpszTypeName,&pRunOpts->ulTypeNum);
						if (pRunOpts->lpszTypeName)
						{
							// Have a name to look up
							pRunOpts->ulTypeNum = PropTypeNameToPropType(pRunOpts->lpszTypeName);
							if (ulNoMatch == pRunOpts->ulTypeNum)
							{
								printf("Property type \"%ws\" not found\n",pRunOpts->lpszTypeName);
							}
						}
						i++;
					}
					break;
				// Smart View parsing
				case 'p':
					if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
					if (i+1 < argc)
					{
						LPSTR szEndPtr = NULL;
						pRunOpts->ulParser = strtoul(argv[i+1],&szEndPtr,10);
						i++;
					}
					break;
				case 'i':
					if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
					if (i+1 < argc)
					{
						EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszInput));
						i++;
					}
					break;
				case 'o':
					if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
					if (i+1 < argc)
					{
						EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszOutput));
						i++;
					}
					break;
				default:
					// display help
					return false;
					break;
				}
			}
			break;
		default:
			// naked option without a flag, must be a property name or number
			if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
			if (pRunOpts->lpszPropName) return false; // He's already got one, you see.
			EC_H(AnsiToUnicode(argv[i],&pRunOpts->lpszPropName));
			break;
		}
	}

	if (pRunOpts->lpszPropName)
	{
		ULONG ulArg = NULL;
		LPWSTR szEndPtr = NULL;
		ulArg = wcstoul(pRunOpts->lpszPropName,&szEndPtr,pRunOpts->bDoDecimal?10:16);

		// if szEndPtr is pointing to something other than NULL, this must be a string
		if (!szEndPtr || *szEndPtr)
		{
			ulArg = NULL;
		}

		pRunOpts->ulPropNum = ulArg;
	}

	// Validate that we have bare minimum to run
	if (cmdmodeUnknown == pRunOpts->Mode) return false;

	if (cmdmodePropTag == pRunOpts->Mode)
	{
		if (pRunOpts->bDoType && !pRunOpts->bDoPartialSearch && (pRunOpts->lpszPropName != NULL)) return false;
	}
	else if (cmdmodeSmartView == pRunOpts->Mode)
	{
		if (!pRunOpts->ulParser || !pRunOpts->lpszInput) return false;
	}

	// Didn't fail - return true
	return true;
} // ParseArgs

void DoPropTags(MYOPTIONS ProgOpts)
{
	if (!ProgOpts.bDoDispid)
	{
		// Canonicalize our prop tag a smidge
		ULONG ulPropID = NULL;
		ULONG ulPropType = NULL;

		if (ProgOpts.ulPropNum & 0xffff0000) // dealing with a full prop tag
		{
			ulPropID   = PROP_ID(ProgOpts.ulPropNum);
			ulPropType = PROP_TYPE(ProgOpts.ulPropNum);
		}
		else
		{
			ulPropID   = ProgOpts.ulPropNum;
			ulPropType = PT_UNSPECIFIED;
		}

		// put everything back together
		ProgOpts.ulPropNum = PROP_TAG(ulPropType,ulPropID);
	}

	if (ProgOpts.bDoPartialSearch)
	{
		if (ProgOpts.bDoType && ulNoMatch == ProgOpts.ulTypeNum)
		{
			DisplayUsage();
		}
		else
		{
			if (ProgOpts.bDoDispid)
			{
				PrintDispIDFromPartialName(ProgOpts.lpszPropName,ProgOpts.ulTypeNum);
			}
			else
			{
				PrintTagFromPartialName(ProgOpts.lpszPropName,ProgOpts.ulTypeNum);
			}
		}
	}
	// If we weren't asked about a property, maybe we were asked about types
	else if (ProgOpts.bDoType)
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
	else if (ProgOpts.bDoDispid)
	{
		if (ProgOpts.ulPropNum)
		{
			PrintDispIDFromNum(ProgOpts.ulPropNum);
		}
		else if (ProgOpts.lpszPropName)
		{
			PrintDispIDFromName(ProgOpts.lpszPropName);
		}
	}
	else
	{
		if (ProgOpts.ulPropNum)
		{
			if (ProgOpts.bDoType && ulNoMatch != ProgOpts.ulTypeNum)
			{
				ProgOpts.ulPropNum = CHANGE_PROP_TYPE(ProgOpts.ulPropNum,ProgOpts.ulTypeNum);
			}
			PrintTagFromNum(ProgOpts.ulPropNum);
		}
		else if (ProgOpts.lpszPropName)
		{
			PrintTagFromName(ProgOpts.lpszPropName);
		}
	}
} // DoPropTags

void DoSmartView(MYOPTIONS ProgOpts)
{
	// Ignore the reg key that disables smart view parsing
	RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD = true;

	ULONG ulStructType = NULL;

	if (ProgOpts.ulParser && ProgOpts.ulParser < g_cuidParsingTypes)
	{
		ulStructType = g_uidParsingTypes[ProgOpts.ulParser];
	}

	if (ulStructType)
	{
		FILE* fIn = NULL;
		FILE* fOut = NULL;
		fIn = _wfopen(ProgOpts.lpszInput,L"rb");
		if (!fIn) printf("Cannot open input file %ws\n",ProgOpts.lpszInput);
		if (ProgOpts.lpszOutput)
		{
			fOut = _wfopen(ProgOpts.lpszOutput,L"wb");
			if (!fOut) printf("Cannot open output file %ws\n",ProgOpts.lpszOutput);
		}

		if (fIn && (!ProgOpts.lpszOutput || fOut))
		{
			int iDesc = _fileno(fIn);
			long iLength = _filelength(iDesc);

			LPTSTR szIn = new TCHAR[iLength+1]; // +1 for NULL
			if (szIn)
			{
				memset(szIn,0,sizeof(TCHAR)*(iLength+1));
				fread(szIn,sizeof(TCHAR),iLength,fIn);
				SBinary Bin = {0};
				
				if (MyBinFromHex(
					(LPCTSTR) szIn,
					NULL,
					&Bin.cb))
				{
					Bin.lpb = new BYTE[Bin.cb];
					if (Bin.lpb)
					{
						if (MyBinFromHex(
							(LPCTSTR) szIn,
							Bin.lpb,
							&Bin.cb))
						{
							LPTSTR szString = NULL;
							InterpretBinaryAsString(Bin,ulStructType,NULL,NULL,&szString);
							if (fOut)
							{
								_fputts(szString,fOut);
							}
							else
							{
								printf("%s\n",szString);
							}
							delete[] szString;
						}
					}
					delete[] Bin.lpb;
				}
				delete[] szIn;
			}
		}

		if (fOut) fclose(fOut);
		if (fIn) fclose(fIn);
	}
} // DoSmartView

void main(int argc, char * argv[])
{
	// Initialize MFC for LoadString support later on
#pragma warning(push)
#pragma warning(disable:6309)
#pragma warning(disable:6387)
   if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0)) return;
#pragma warning(pop)

   LoadAddIns();

	MYOPTIONS ProgOpts;

	if (!ParseArgs(argc, argv, &ProgOpts))
	{
		DisplayUsage();
		delete[] ProgOpts.lpszPropName;
		delete[] ProgOpts.lpszTypeName;
		return;
	}

	switch (ProgOpts.Mode)
	{
	case cmdmodePropTag:
		DoPropTags(ProgOpts);
		break;
	case cmdmodeSmartView:
		DoSmartView(ProgOpts);
		break;
	}

	delete[] ProgOpts.lpszPropName;
	delete[] ProgOpts.lpszTypeName;
} // main