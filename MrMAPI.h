#pragma once

#define ulNoMatch 0xffffffff

enum CmdMode {
	cmdmodeUnknown = 0,
	cmdmodePropTag,
	cmdmodeSmartView,
};

struct MYOPTIONS
{
	CmdMode Mode;
	BOOL  bDoPartialSearch;
	BOOL  bDoType;
	BOOL  bDoDispid;
	BOOL  bDoDecimal;
	ULONG ulPropNum;
	LPWSTR lpszPropName;
	ULONG ulTypeNum;
	LPWSTR lpszTypeName;
	ULONG ulParser;
	LPWSTR lpszInput;
	LPWSTR lpszOutput;
};

////////////////////////////////////////////////////////////////////////////////
// Array lookups
////////////////////////////////////////////////////////////////////////////////

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
						 ULONG* lpulFirstPartial);

// Searches a NAMEID_ARRAY_ENTRY array for a target dispid.
// Exact matches are those that match
// If no hits, then ulNoMatch should be returned for lpulFirstExact
void FindNameIDArrayMatches(LONG lTarget,
							NAMEID_ARRAY_ENTRY* MyArray,
							ULONG ulMyArray,
							ULONG* lpulNumExacts,
							ULONG* lpulFirstExact);


////////////////////////////////////////////////////////////////////////////////
// Pretty printing functions
////////////////////////////////////////////////////////////////////////////////

// prints the type of a prop tag
// no pretty stuff or \n - calling function gets to do that
void PrintType(ULONG ulPropTag);

void PrintKnownTypes();

// Print the tag found in the array at ulRow
void PrintTag(ULONG ulRow);

void PrintTagFromNum(ULONG ulPropTag);

void PrintTagFromName(LPCWSTR lpszPropName);

// Search for properties matching lpszPropName on a substring
// If ulType isn't ulNoMatch, restrict on the property type as well
void PrintTagFromPartialName(LPCWSTR lpszPropName, ULONG ulType);

void PrintGUID(LPCGUID lpGUID);

void PrintDispID(ULONG ulRow);

void PrintDispIDFromNum(ULONG ulDispID);

void PrintDispIDFromName(LPCWSTR lpszDispIDName);

// Search for properties matching lpszPropName on a substring
void PrintDispIDFromPartialName(LPCWSTR lpszDispIDName, ULONG ulType);

////////////////////////////////////////////////////////////////////////////////
// Command line interface functions
////////////////////////////////////////////////////////////////////////////////

void DisplayUsage();

// scans an arg and returns the string or hex number that it represents
void GetArg(char* szArgIn, WCHAR** lpszArgOut, ULONG* ulArgOut);

// Parses command line arguments and fills out MYOPTIONS
BOOL ParseArgs(int argc, char * argv[], MYOPTIONS * pRunOpts);

void main(int argc, char * argv[]);
////////////////////////////////////////////////////////////////////////////////