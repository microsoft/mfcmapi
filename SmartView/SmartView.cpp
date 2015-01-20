#include "stdafx.h"
#include "..\stdafx.h"
#include "SmartView.h"
#include "BinaryParser.h"
#include "..\InterpretProp2.h"
#include "..\InterpretProp.h"
#include "..\ExtraPropTags.h"
#include "..\MAPIFunctions.h"
#include "..\String.h"
#include "..\guids.h"
#include "..\MySecInfo.h"
#include "..\NamedPropCache.h"
#include "..\ParseProperty.h"

#include "SmartViewParser.h"
#include "PCL.h"
#include "TombStone.h"
#include "VerbStream.h"
#include "NickNameCache.h"
#include "FolderUserFieldStream.h"
#include "RecipientRowStream.h"
#include "WebViewPersistStream.h"
#include "FlatEntryList.h"
#include "AdditionalRenEntryIDs.h"
#include "PropertyDefinitionStream.h"
#include "SearchFolderDefinition.h"

#define _MaxBytes 0xFFFF
#define _MaxDepth 50
#define _MaxEID 500
#define _MaxEntriesSmall 500
#define _MaxEntriesLarge 1000
#define _MaxEntriesExtraLarge 1500
#define _MaxEntriesEnormous 10000

ULONG InterpretMVLongAsString(SLongArray myLongArray, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_ LPGUID lpguidNamedProp, _Deref_out_opt_z_ LPWSTR* lpszResultString);

// Functions to parse PT_LONG/PT-I2 properties

_Check_return_ CString RTimeToString(DWORD rTime);
_Check_return_ LPWSTR RTimeToSzString(DWORD rTime, bool bLabel);
_Check_return_ LPWSTR PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel);

// End: Functions to parse PT_LONG/PT-I2 properties

// After 'No Parsing', these are in alphabetical order
NAME_ARRAY_ENTRY g_uidParsingTypes[] = {
	{ IDS_STNOPARSING, L"Choose Smart View Parser" }, // STRING_OK
	{ IDS_STADDITIONALRENENTRYIDSEX, L"Additional Ren Entry IDs Ex" }, // STRING_OK
	{ IDS_STAPPOINTMENTRECURRENCEPATTERN, L"Appointment Recurrence Pattern" }, // STRING_OK
	{ IDS_STCONVERSATIONINDEX, L"Conversation Index" }, // STRING_OK
	{ IDS_STENTRYID, L"Entry Id" }, // STRING_OK
	{ IDS_STENTRYLIST, L"Entry List" }, // STRING_OK
	{ IDS_STEXTENDEDFOLDERFLAGS, L"Extended Folder Flags" }, // STRING_OK
	{ IDS_STEXTENDEDRULECONDITION, L"Extended Rule Condition" }, // STRING_OK
	{ IDS_STFLATENTRYLIST, L"Flat Entry List" }, // STRING_OK
	{ IDS_STFOLDERUSERFIELDS, L"Folder User Fields Stream" }, // STRING_OK
	{ IDS_STGLOBALOBJECTID, L"Global Object Id" }, // STRING_OK
	{ IDS_STPROPERTY, L"Property" }, // STRING_OK
	{ IDS_STPROPERTYDEFINITIONSTREAM, L"Property Definition Stream" }, // STRING_OK
	{ IDS_STRECIPIENTROWSTREAM, L"Recipient Row Stream" }, // STRING_OK
	{ IDS_STRECURRENCEPATTERN, L"Recurrence Pattern" }, // STRING_OK
	{ IDS_STREPORTTAG, L"Report Tag" }, // STRING_OK
	{ IDS_STRESTRICTION, L"Restriction" }, // STRING_OK
	{ IDS_STRULECONDITION, L"Rule Condition" }, // STRING_OK
	{ IDS_STSEARCHFOLDERDEFINITION, L"Search Folder Definition" }, // STRING_OK
	{ IDS_STSECURITYDESCRIPTOR, L"Security Descriptor" }, // STRING_OK
	{ IDS_STSID, L"SID" }, // STRING_OK
	{ IDS_STTASKASSIGNERS, L"Task Assigners" }, // STRING_OK
	{ IDS_STTIMEZONE, L"Time Zone" }, // STRING_OK
	{ IDS_STTIMEZONEDEFINITION, L"Time Zone Definition" }, // STRING_OK
	{ IDS_STWEBVIEWPERSISTSTREAM, L"Web View Persistence Object Stream" }, // STRING_OK
	{ IDS_STNICKNAMECACHE, L"Nickname Cache" }, // STRING_OK
	{ IDS_STENCODEENTRYID, L"Encode Entry ID" }, // STRING_OK
	{ IDS_STDECODEENTRYID, L"Decode Entry ID" }, // STRING_OK
	{ IDS_STVERBSTREAM, L"Verb Stream" }, // STRING_OK
	{ IDS_STTOMBSTONE, L"Tombstone" }, // STRING_OK
	{ IDS_STPCL, L"Predecessor Change List" }, // STRING_OK
};
ULONG g_cuidParsingTypes = _countof(g_uidParsingTypes);

typedef LPVOID BINTOSTRUCT(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
typedef BINTOSTRUCT *LPBINTOSTRUCT;
typedef void DELETESTRUCT(LPVOID lpStruct);
typedef DELETESTRUCT *LPDELETESTRUCT;
typedef LPWSTR STRUCTTOSTRING(LPVOID lpStruct);
typedef STRUCTTOSTRING *LPSTRUCTTOSTRING;

typedef void (BINTOSTRING)(SBinary myBin, LPWSTR* lpszResultString);
typedef BINTOSTRING *LPBINTOSTRING;
struct SMART_VIEW_PARSERS_ENTRY
{
	ULONG iStructType;
	LPBINTOSTRUCT fBinToStruct;
	LPDELETESTRUCT fDeleteStruct;
	LPSTRUCTTOSTRING fStructToString;
};
#define MAKE_SV_ENTRY(_fIDSType, _fType) {(_fIDSType), (LPBINTOSTRUCT) BinTo##_fType, (LPDELETESTRUCT) Delete##_fType, (LPSTRUCTTOSTRING) _fType##ToString},
SMART_VIEW_PARSERS_ENTRY g_SmartViewParsers[] = {
	{ IDS_STNOPARSING, NULL, NULL, NULL },
	MAKE_SV_ENTRY(IDS_STTIMEZONEDEFINITION, TimeZoneDefinitionStruct)
	MAKE_SV_ENTRY(IDS_STTIMEZONE, TimeZoneStruct)
	// MAKE_SV_ENTRY(IDS_STSECURITYDESCRIPTOR, NULL}
	MAKE_SV_ENTRY(IDS_STEXTENDEDFOLDERFLAGS, ExtendedFlagsStruct)
	MAKE_SV_ENTRY(IDS_STAPPOINTMENTRECURRENCEPATTERN, AppointmentRecurrencePatternStruct)
	MAKE_SV_ENTRY(IDS_STRECURRENCEPATTERN, RecurrencePatternStruct)
	MAKE_SV_ENTRY(IDS_STREPORTTAG, ReportTagStruct)
	MAKE_SV_ENTRY(IDS_STCONVERSATIONINDEX, ConversationIndexStruct)
	MAKE_SV_ENTRY(IDS_STTASKASSIGNERS, TaskAssignersStruct)
	MAKE_SV_ENTRY(IDS_STGLOBALOBJECTID, GlobalObjectIdStruct)
	MAKE_SV_ENTRY(IDS_STENTRYID, EntryIdStruct)
	MAKE_SV_ENTRY(IDS_STENTRYLIST, EntryListStruct)
	MAKE_SV_ENTRY(IDS_STPROPERTY, PropertyStruct)
	MAKE_SV_ENTRY(IDS_STRESTRICTION, RestrictionStruct)
	// MAKE_SV_ENTRY(IDS_STRULECONDITION, RuleConditionStruct)
	// MAKE_SV_ENTRY(IDS_STEXTENDEDRULECONDITION, RuleConditionStruct)
	// MAKE_SV_ENTRY(IDS_STSID, SIDStruct)
	// MAKE_SV_ENTRY(IDS_STDECODEENTRYID)
	// MAKE_SV_ENTRY(IDS_STENCODEENTRYID)
};
ULONG g_cSmartViewParsers = _countof(g_SmartViewParsers);

LPSMARTVIEWPARSER GetSmartViewParser(DWORD_PTR iStructType, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	switch (iStructType)
	{
	case IDS_STNOPARSING:
		return NULL;
		break;
	case IDS_STTOMBSTONE:
		return new TombStone(cbBin, lpBin);
		break;
	case IDS_STPCL:
		return new PCL(cbBin, lpBin);
		break;
	case IDS_STVERBSTREAM:
		return new VerbStream(cbBin, lpBin);
		break;
	case IDS_STNICKNAMECACHE:
		return new NickNameCache(cbBin, lpBin);
		break;
	case IDS_STFOLDERUSERFIELDS:
		return new FolderUserFieldStream(cbBin, lpBin);
		break;
	case IDS_STRECIPIENTROWSTREAM:
		return new RecipientRowStream(cbBin, lpBin);
		break;
	case IDS_STWEBVIEWPERSISTSTREAM:
		return new WebViewPersistStream(cbBin, lpBin);
		break;
	case IDS_STFLATENTRYLIST:
		return new FlatEntryList(cbBin, lpBin);
		break;
	case IDS_STADDITIONALRENENTRYIDSEX:
		return new AdditionalRenEntryIDs(cbBin, lpBin);
		break;
	case IDS_STPROPERTYDEFINITIONSTREAM:
		return new PropertyDefinitionStream(cbBin, lpBin);
		break;
	case IDS_STSEARCHFOLDERDEFINITION:
		return new SearchFolderDefinition(cbBin, lpBin);
		break;
	}

	return NULL;
}

SMARTVIEW_PARSER_ARRAY_ENTRY g_NumStructArray[] =
{
	BINARY_STRUCTURE_ENTRY(PR_RULE_ID, IDS_STPTI8)
	BINARY_STRUCTURE_ENTRY(PidTagFolderId, IDS_STSFIDMID)
	BINARY_STRUCTURE_ENTRY(PidTagMid, IDS_STSFIDMID)
	BINARY_STRUCTURE_ENTRY(PR_WB_SF_LAST_USED, IDS_STLONGRTIME)
	BINARY_STRUCTURE_ENTRY(PR_WB_SF_EXPIRATION, IDS_STLONGRTIME)
	BINARY_STRUCTURE_ENTRY(PR_FREEBUSY_PUBLISH_START, IDS_STLONGRTIME)
	BINARY_STRUCTURE_ENTRY(PR_FREEBUSY_PUBLISH_END, IDS_STLONGRTIME)
};

LPSMARTVIEW_PARSER_ARRAY_ENTRY NumStructArray = g_NumStructArray;
ULONG ulNumStructArray = _countof(g_NumStructArray);

_Check_return_ ULONG BuildFlagIndexFromTag(ULONG ulPropTag,
	ULONG ulPropNameID,
	_In_opt_z_ LPWSTR lpszPropNameString,
	_In_opt_ LPCGUID lpguidNamedProp)
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
		if (*lpguidNamedProp == PSETID_Meeting)        ulGuid = guidPSETID_Meeting;
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
			return PROP_TAG(ulGuid, ulPropNameID);
		}
		// Case not handled yet
		// else if (ulGuid && lpszPropNameString)
		// {
		// }
	}
	return NULL;
} // BuildFlagIndexFromTag

_Check_return_ ULONG FindSmartViewParserForProp(_In_count_(ulParserArray) LPSMARTVIEW_PARSER_ARRAY_ENTRY lpParserArray, ULONG ulParserArray, const ULONG ulPropTag, const ULONG ulPropNameID, _In_opt_ const LPCGUID lpguidNamedProp)
{
	if (!lpParserArray) return 0;
	ULONG	ulCurEntry = 0;
	ULONG	ulIndex = BuildFlagIndexFromTag(ulPropTag, ulPropNameID, NULL, lpguidNamedProp);
	bool	bMV = (PROP_TYPE(ulPropTag) & MV_FLAG) == MV_FLAG;

	while (ulCurEntry < ulParserArray)
	{
		if (lpParserArray[ulCurEntry].ulIndex == ulIndex &&
			lpParserArray[ulCurEntry].bMV == bMV)
			return lpParserArray[ulCurEntry].iStructType;
		ulCurEntry++;
	}

	return 0;
} // FindSmartViewParserForProp

// lpszSmartView allocated with new, delete with delete[]
ULONG InterpretPropSmartView(_In_ LPSPropValue lpProp, // required property value
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	bool bMVRow, // did the row come from a MV prop?
	_Deref_out_opt_z_ LPWSTR* lpszSmartView) // Built from lpProp & lpMAPIProp
{
	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return NULL;
	if (!lpszSmartView) return NULL;
	*lpszSmartView = NULL;
	if (!lpProp) return NULL;

	HRESULT hRes = S_OK;
	ULONG iStructType = NULL;

	// Named Props
	LPMAPINAMEID* lppPropNames = 0;

	// If we weren't passed named property information and we need it, look it up
	// We check bIsAB here - some address book providers return garbage which will crash us
	if (!lpNameID &&
		lpMAPIProp && // if we have an object
		!bIsAB &&
		RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD && // and we're parsing named props
		(RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD || PROP_ID(lpProp->ulPropTag) >= 0x8000)) // and it's either a named prop or we're doing all props
	{
		SPropTagArray	tag = { 0 };
		LPSPropTagArray	lpTag = &tag;
		ULONG			ulPropNames = 0;
		tag.cValues = 1;
		tag.aulPropTag[0] = lpProp->ulPropTag;

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

	ULONG ulPropNameID = NULL;
	LPGUID lpPropNameGUID = NULL;

	if (lpNameID)
	{
		lpPropNameGUID = lpNameID->lpguid;
		if (lpNameID->ulKind == MNID_ID)
		{
			ulPropNameID = lpNameID->Kind.lID;
		}
	}

	switch (PROP_TYPE(lpProp->ulPropTag))
	{
	case PT_LONG:
	case PT_I2:
	case PT_I8:
	{
				  iStructType = InterpretNumberAsString(lpProp->Value, lpProp->ulPropTag, ulPropNameID, NULL, lpPropNameGUID, true, lpszSmartView);
	}
		break;
	case PT_MV_LONG:
	{
					   iStructType = InterpretMVLongAsString(lpProp->Value.MVl, lpProp->ulPropTag, ulPropNameID, lpPropNameGUID, lpszSmartView);
	}
		break;
	case PT_BINARY:
	{
					  ULONG ulLookupPropTag = lpProp->ulPropTag;
					  if (bMVRow) ulLookupPropTag |= MV_FLAG;

					  iStructType = FindSmartViewParserForProp(SmartViewParserArray, ulSmartViewParserArray, ulLookupPropTag, ulPropNameID, lpPropNameGUID);
					  // We special-case this property
					  if (!iStructType && PR_ROAMING_BINARYSTREAM == ulLookupPropTag && lpMAPIProp)
					  {
						  LPSPropValue lpPropSubject = NULL;

						  WC_MAPI(HrGetOneProp(
							  lpMAPIProp,
							  PR_SUBJECT_W,
							  &lpPropSubject));

						  if (CheckStringProp(lpPropSubject, PT_UNICODE) && 0 == wcscmp(lpPropSubject->Value.lpszW, L"IPM.Configuration.Autocomplete"))
						  {
							  iStructType = IDS_STNICKNAMECACHE;
						  }
					  }

					  if (iStructType)
					  {
						  InterpretBinaryAsString(lpProp->Value.bin, iStructType, lpMAPIProp, lpProp->ulPropTag, lpszSmartView);
					  }
	}
		break;
	case PT_MV_BINARY:
	{
						 iStructType = FindSmartViewParserForProp(SmartViewParserArray, ulSmartViewParserArray, lpProp->ulPropTag, ulPropNameID, lpPropNameGUID);
						 if (iStructType)
						 {
							 InterpretMVBinaryAsString(lpProp->Value.MVbin, iStructType, lpMAPIProp, lpProp->ulPropTag, lpszSmartView);
						 }
	}
		break;
	}
	MAPIFreeBuffer(lppPropNames);

	return iStructType;
} // InterpretPropSmartView

void InterpretMVBinaryAsString(SBinaryArray myBinArray, DWORD_PTR iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_z_ LPWSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;

	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return;

	ULONG ulRow = 0;
	CString szResult;
	CString szTmp;
	LPWSTR szSmartView = NULL;

	for (ulRow = 0; ulRow < myBinArray.cValues; ulRow++)
	{
		if (ulRow != 0)
		{
			szResult += _T("\r\n\r\n"); // STRING_OK
		}
		szTmp.FormatMessage(IDS_MVROWBIN,
			ulRow);
		szResult += szTmp;
		InterpretBinaryAsString(myBinArray.lpbin[ulRow], iStructType, lpMAPIProp, ulPropTag, &szSmartView);
		if (szSmartView)
		{
			szResult += szSmartView;
			delete[] szSmartView;
			szSmartView = NULL;
		}
	}

	*lpszResultString = CStringToLPWSTR(szResult);
} // InterpretMVBinaryAsString

ULONG InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag, _Deref_out_opt_z_ LPWSTR* lpszResultString)
{
	_PV pV = { 0 };
	pV.ul = ulVal;
	return InterpretNumberAsString(pV, ulPropTag, NULL, NULL, NULL, false, lpszResultString);
} // InterpretNumberAsStringProp

ULONG InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp, _Deref_out_opt_z_ LPWSTR* lpszResultString)
{
	_PV pV = { 0 };
	pV.ul = ulVal;
	return InterpretNumberAsString(pV, PT_LONG, ulPropNameID, NULL, lpguidNamedProp, false, lpszResultString);
} // InterpretNumberAsStringNamedProp

// Interprets a PT_LONG, PT_I2. or PT_I8 found in lpProp and returns a string allocated with new
// Free the string with delete[]
// Will not return a string if the lpProp is not a PT_LONG/PT_I2/PT_I8 or we don't recognize the property
// Will use named property details to look up named property flags
ULONG InterpretNumberAsString(_PV pV, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_z_ LPWSTR lpszPropNameString, _In_opt_ LPCGUID lpguidNamedProp, bool bLabel, _Deref_out_opt_z_ LPWSTR* lpszResultString)
{
	if (lpszResultString) *lpszResultString = NULL;
	if (!ulPropTag || !lpszResultString)
	{
		return NULL;
	}
	if (PROP_TYPE(ulPropTag) != PT_LONG &&
		PROP_TYPE(ulPropTag) != PT_I2 &&
		PROP_TYPE(ulPropTag) != PT_I8)
	{
		return NULL;
	}

	ULONG iParser = FindSmartViewParserForProp(NumStructArray, ulNumStructArray, ulPropTag, ulPropNameID, lpguidNamedProp);
	switch (iParser)
	{
	case IDS_STLONGRTIME:
		*lpszResultString = RTimeToSzString(pV.ul, bLabel);
		break;
	case IDS_STPTI8:
		*lpszResultString = PTI8ToSzString(pV.li, bLabel);
		break;
	case IDS_STSFIDMID:
		*lpszResultString = FidMidToSzString(pV.li.QuadPart, bLabel);
		break;
		// insert future parsers here
	default:
	{
			   ULONG ulPropID = BuildFlagIndexFromTag(ulPropTag, ulPropNameID, lpszPropNameString, lpguidNamedProp);
			   if (ulPropID)
			   {
				   CString szPrefix;
				   HRESULT hRes = S_OK;
				   if (bLabel)
				   {
					   EC_B(szPrefix.LoadString(IDS_FLAGS_PREFIX));
				   }

#ifdef UNICODE
				   InterpretFlags(ulPropID, pV.ul, szPrefix, lpszResultString);
#else
				   LPSTR lpszResultStringA = NULL;
				   LPWSTR lpszResultStringW = NULL;
				   InterpretFlags(ulPropID, pV.ul, szPrefix, &lpszResultStringA);
				   EC_H(AnsiToUnicode(lpszResultStringA, &lpszResultStringW));
				   *lpszResultString = lpszResultStringW;
				   delete[] lpszResultStringA;
#endif
			   }
	}
		break;
	}

	return iParser;
} // InterpretNumberAsString

ULONG InterpretMVLongAsString(SLongArray myLongArray, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_ LPGUID lpguidNamedProp, _Deref_out_opt_z_ LPWSTR* lpszResultString)
{
	if (!lpszResultString) return NULL;
	*lpszResultString = NULL;

	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return NULL;

	ULONG iStructType = NULL;
	ULONG ulRow = 0;
	CString szResult;
	CString szTmp;
	LPWSTR szSmartView = NULL;
	bool bHasData = false;

	for (ulRow = 0; ulRow < myLongArray.cValues; ulRow++)
	{
		_PV pV = { 0 };
		pV.ul = myLongArray.lpl[ulRow];
		iStructType = InterpretNumberAsString(pV, CHANGE_PROP_TYPE(ulPropTag, PT_LONG), ulPropNameID, NULL, lpguidNamedProp, true, &szSmartView);
		if (szSmartView)
		{
			if (bHasData)
			{
				szResult += _T("\r\n"); // STRING_OK
			}
			bHasData = true;
			szTmp.FormatMessage(IDS_MVROWLONG,
				ulRow,
				szSmartView ? szSmartView : L"");
			szResult += szTmp;
			delete[] szSmartView;
			szSmartView = NULL;
		}
	}

	*lpszResultString = CStringToLPWSTR(szResult);

	return iStructType;
} // InterpretMVLongAsString

// lpszResultString allocated with new, delete with delete[]
void InterpretBinaryAsString(SBinary myBin, DWORD_PTR iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_z_ LPWSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;

	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return;
	LPWSTR szResultString = NULL;

	ULONG i = 0;
	bool bParsed = false;
	for (i = 0; i < g_cSmartViewParsers; i++)
	{
		if (iStructType == g_SmartViewParsers[i].iStructType)
		{
			if (g_SmartViewParsers[i].fBinToStruct && g_SmartViewParsers[i].fStructToString && g_SmartViewParsers[i].fDeleteStruct)
			{
				LPVOID pStruct = g_SmartViewParsers[i].fBinToStruct(myBin.cb, myBin.lpb);
				if (pStruct)
				{
					szResultString = g_SmartViewParsers[i].fStructToString(pStruct);
					g_SmartViewParsers[i].fDeleteStruct(pStruct);
					pStruct = NULL;
				}
			}

			bParsed = true;
			break;
		}
	}

	if (!bParsed)
	{
		LPSMARTVIEWPARSER svp = GetSmartViewParser(iStructType, myBin.cb, myBin.lpb);
		if (svp)
		{
			szResultString = svp->ToString();
			delete svp;
			svp = NULL;
			bParsed = true;
		}
	}

	// These parsers have some special casing
	if (!bParsed)
	{
		switch (iStructType)
		{
		case IDS_STSECURITYDESCRIPTOR:
			SDBinToString(myBin, lpMAPIProp, ulPropTag, &szResultString);
			break;
		case IDS_STRULECONDITION:
			RuleConditionToString(myBin, &szResultString, false);
			break;
		case IDS_STEXTENDEDRULECONDITION:
			RuleConditionToString(myBin, &szResultString, true);
			break;
		case IDS_STSID:
			SIDBinToString(myBin, &szResultString);
			break;
		case IDS_STDECODEENTRYID:
			szResultString = DecodeID(myBin.cb, myBin.lpb);
			break;
		case IDS_STENCODEENTRYID:
			szResultString = EncodeID(myBin.cb, (LPENTRYID)myBin.lpb);
			break;
		}
	}

	*lpszResultString = szResultString;
}

_Check_return_ CString RTimeToString(DWORD rTime)
{
	CString PropString;
	FILETIME fTime = { 0 };
	LARGE_INTEGER liNumSec = { 0 };
	liNumSec.LowPart = rTime;
	// Resolution of RTime is in minutes, FILETIME is in 100 nanosecond intervals
	// Scale between the two is 10000000*60
	liNumSec.QuadPart = liNumSec.QuadPart * 10000000 * 60;
	fTime.dwLowDateTime = liNumSec.LowPart;
	fTime.dwHighDateTime = liNumSec.HighPart;
	FileTimeToString(&fTime, &PropString, NULL);
	return PropString;
} // RTimeToString

_Check_return_ LPWSTR RTimeToSzString(DWORD rTime, bool bLabel)
{
	CString szRTime;
	if (bLabel)
	{
		szRTime.FormatMessage(_T("RTime: ")); // STRING_OK
	}
	szRTime += RTimeToString(rTime);
	return CStringToLPWSTR(szRTime);
} // RTimeToSzString

_Check_return_ LPWSTR PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel)
{
	CString szI8;
	if (bLabel)
	{
		szI8.FormatMessage(IDS_PTI8FORMATLABEL, liI8.LowPart, liI8.HighPart);
	}
	else
	{
		szI8.FormatMessage(IDS_PTI8FORMAT, liI8.LowPart, liI8.HighPart);
	}
	return CStringToLPWSTR(szI8);
} // PTI8ToSzString

typedef WORD REPLID;
#define cbGlobcnt 6

struct ID
{
	REPLID replid;
	BYTE globcnt[cbGlobcnt];
};

_Check_return_ WORD WGetReplId(ID id)
{
	return id.replid;
} // WGetReplId

_Check_return_ ULONGLONG UllGetIdGlobcnt(ID id)
{
	ULONGLONG ul = 0;
	for (int i = 0; i < cbGlobcnt; ++i)
	{
		ul <<= 8;
		ul += id.globcnt[i];
	}

	return ul;
} // UllGetIdGlobcnt

_Check_return_ LPWSTR FidMidToSzString(LONGLONG llID, bool bLabel)
{
	CString szFidMid;
	ID* pid = (ID*)&llID;
	if (bLabel)
	{
		szFidMid.FormatMessage(IDS_FIDMIDFORMATLABEL, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
	}
	else
	{
		szFidMid.FormatMessage(IDS_FIDMIDFORMAT, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
	}
	return CStringToLPWSTR(szFidMid);
} // FidMidToSzString

_Check_return_ CString JunkDataToString(size_t cbJunkData, _In_count_(cbJunkData) LPBYTE lpJunkData)
{
	if (!cbJunkData || !lpJunkData) return _T("");
	DebugPrint(DBGSmartView, _T("Had 0x%08X = %u bytes left over.\n"), (int)cbJunkData, (UINT)cbJunkData);
	CString szTmp;
	SBinary sBin = { 0 };

	sBin.cb = (ULONG)cbJunkData;
	sBin.lpb = lpJunkData;
	szTmp.FormatMessage(IDS_JUNKDATASIZE,
		cbJunkData);
	szTmp += BinToHexString(&sBin, true).c_str();
	return szTmp;
} // JunkDataToString

//////////////////////////////////////////////////////////////////////////
// RecurrencePatternStruct
//////////////////////////////////////////////////////////////////////////

_Check_return_ RecurrencePatternStruct* BinToRecurrencePatternStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	return BinToRecurrencePatternStructWithSize(cbBin, lpBin, NULL);
} // BinToRecurrencePatternStruct

// Parses lpBin as a RecurrencePatternStruct
// lpcbBytesRead returns the number of bytes consumed
// Allocates return value with new.
// clean up with delete.
_Check_return_ RecurrencePatternStruct* BinToRecurrencePatternStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead)
{
	if (!lpBin) return NULL;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	CBinaryParser Parser(cbBin, lpBin);

	RecurrencePatternStruct rpPattern = { 0 };

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

	if (rpPattern.DeletedInstanceCount && rpPattern.DeletedInstanceCount < _MaxEntriesSmall)
	{
		rpPattern.DeletedInstanceDates = new DWORD[rpPattern.DeletedInstanceCount];
		if (rpPattern.DeletedInstanceDates)
		{
			memset(rpPattern.DeletedInstanceDates, 0, sizeof(DWORD)* rpPattern.DeletedInstanceCount);
			DWORD i = 0;
			for (i = 0; i < rpPattern.DeletedInstanceCount; i++)
			{
				Parser.GetDWORD(&rpPattern.DeletedInstanceDates[i]);
			}
		}
	}

	Parser.GetDWORD(&rpPattern.ModifiedInstanceCount);

	if (rpPattern.ModifiedInstanceCount &&
		rpPattern.ModifiedInstanceCount <= rpPattern.DeletedInstanceCount &&
		rpPattern.ModifiedInstanceCount < _MaxEntriesSmall)
	{
		rpPattern.ModifiedInstanceDates = new DWORD[rpPattern.ModifiedInstanceCount];
		if (rpPattern.ModifiedInstanceDates)
		{
			memset(rpPattern.ModifiedInstanceDates, 0, sizeof(DWORD)* rpPattern.ModifiedInstanceCount);
			DWORD i = 0;
			for (i = 0; i < rpPattern.ModifiedInstanceCount; i++)
			{
				Parser.GetDWORD(&rpPattern.ModifiedInstanceDates[i]);
			}
		}
	}
	Parser.GetDWORD(&rpPattern.StartDate);
	Parser.GetDWORD(&rpPattern.EndDate);

	// Only fill out junk data if we've not been asked to report back how many bytes we read
	// If we've been asked to report back, then someone else will handle the remaining data
	if (!lpcbBytesRead)
	{
		rpPattern.JunkDataSize = Parser.GetRemainingData(&rpPattern.JunkData);
	}
	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();

	RecurrencePatternStruct* prpPattern = new RecurrencePatternStruct;
	if (prpPattern)
	{
		*prpPattern = rpPattern;
	}

	return prpPattern;
} // BinToRecurrencePatternStructWithSize

void DeleteRecurrencePatternStruct(_In_ RecurrencePatternStruct* prpPattern)
{
	if (!prpPattern) return;
	delete[] prpPattern->DeletedInstanceDates;
	delete[] prpPattern->ModifiedInstanceDates;
	delete[] prpPattern->JunkData;
	delete prpPattern;
} // DeleteRecurrencePatternStruct

// result allocated with new
// clean up with delete[]
_Check_return_ LPWSTR RecurrencePatternStructToString(_In_ RecurrencePatternStruct* prpPattern)
{
	if (!prpPattern) return NULL;

	CString szRP;
	CString szTmp;

	LPTSTR szRecurFrequency = NULL;
	LPTSTR szPatternType = NULL;
	LPTSTR szCalendarType = NULL;
	InterpretFlags(flagRecurFrequency, prpPattern->RecurFrequency, &szRecurFrequency);
	InterpretFlags(flagPatternType, prpPattern->PatternType, &szPatternType);
	InterpretFlags(flagCalendarType, prpPattern->CalendarType, &szCalendarType);
	szTmp.FormatMessage(IDS_RPHEADER,
		prpPattern->ReaderVersion,
		prpPattern->WriterVersion,
		prpPattern->RecurFrequency, szRecurFrequency,
		prpPattern->PatternType, szPatternType,
		prpPattern->CalendarType, szCalendarType,
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
		InterpretFlags(flagDOW, prpPattern->PatternTypeSpecific.WeekRecurrencePattern, &szDOW);
		szTmp.FormatMessage(IDS_RPPATTERNWEEK,
			prpPattern->PatternTypeSpecific.WeekRecurrencePattern, szDOW);
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
		InterpretFlags(flagDOW, prpPattern->PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek, &szDOW);
		InterpretFlags(flagN, prpPattern->PatternTypeSpecific.MonthNthRecurrencePattern.N, &szN);
		szTmp.FormatMessage(IDS_RPPATTERNMONTHNTH,
			prpPattern->PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek, szDOW,
			prpPattern->PatternTypeSpecific.MonthNthRecurrencePattern.N, szN);
		szRP += szTmp;
		break;
	}
	delete[] szDOW;
	delete[] szN;
	szDOW = NULL;
	szN = NULL;

	LPTSTR szEndType = NULL;
	LPTSTR szFirstDOW = NULL;
	InterpretFlags(flagEndType, prpPattern->EndType, &szEndType);
	InterpretFlags(flagFirstDOW, prpPattern->FirstDOW, &szFirstDOW);

	szTmp.FormatMessage(IDS_RPHEADER2,
		prpPattern->EndType, szEndType,
		prpPattern->OccurrenceCount,
		prpPattern->FirstDOW, szFirstDOW,
		prpPattern->DeletedInstanceCount);
	szRP += szTmp;
	delete[] szEndType;
	delete[] szFirstDOW;
	szEndType = NULL;
	szFirstDOW = NULL;

	if (prpPattern->DeletedInstanceCount && prpPattern->DeletedInstanceDates)
	{
		DWORD i = 0;
		for (i = 0; i < prpPattern->DeletedInstanceCount; i++)
		{
			szTmp.FormatMessage(IDS_RPDELETEDINSTANCEDATES,
				i, prpPattern->DeletedInstanceDates[i], RTimeToString(prpPattern->DeletedInstanceDates[i]));
			szRP += szTmp;
		}
	}

	szTmp.FormatMessage(IDS_RPMODIFIEDINSTANCECOUNT,
		prpPattern->ModifiedInstanceCount);
	szRP += szTmp;

	if (prpPattern->ModifiedInstanceCount && prpPattern->ModifiedInstanceDates)
	{
		DWORD i = 0;
		for (i = 0; i < prpPattern->ModifiedInstanceCount; i++)
		{
			szTmp.FormatMessage(IDS_RPMODIFIEDINSTANCEDATES,
				i, prpPattern->ModifiedInstanceDates[i], RTimeToString(prpPattern->ModifiedInstanceDates[i]));
			szRP += szTmp;
		}
	}

	szTmp.FormatMessage(IDS_RPDATE,
		prpPattern->StartDate, RTimeToString(prpPattern->StartDate),
		prpPattern->EndDate, RTimeToString(prpPattern->EndDate));
	szRP += szTmp;

	szRP += JunkDataToString(prpPattern->JunkDataSize, prpPattern->JunkData);

	return CStringToLPWSTR(szRP);
} // RecurrencePatternStructToString

//////////////////////////////////////////////////////////////////////////
// End RecurrencePatternStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// AppointmentRecurrencePatternStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new.
// Clean up with DeleteAppointmentRecurrencePatternStruct.
_Check_return_ AppointmentRecurrencePatternStruct* BinToAppointmentRecurrencePatternStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	AppointmentRecurrencePatternStruct arpPattern = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

	size_t cbBinRead = 0;
	arpPattern.RecurrencePattern = BinToRecurrencePatternStructWithSize(cbBin, lpBin, &cbBinRead);
	Parser.Advance(cbBinRead);
	Parser.GetDWORD(&arpPattern.ReaderVersion2);
	Parser.GetDWORD(&arpPattern.WriterVersion2);
	Parser.GetDWORD(&arpPattern.StartTimeOffset);
	Parser.GetDWORD(&arpPattern.EndTimeOffset);
	Parser.GetWORD(&arpPattern.ExceptionCount);

	if (arpPattern.ExceptionCount &&
		arpPattern.ExceptionCount == arpPattern.RecurrencePattern->ModifiedInstanceCount &&
		arpPattern.ExceptionCount < _MaxEntriesSmall)
	{
		arpPattern.ExceptionInfo = new ExceptionInfoStruct[arpPattern.ExceptionCount];
		if (arpPattern.ExceptionInfo)
		{
			memset(arpPattern.ExceptionInfo, 0, sizeof(ExceptionInfoStruct)* arpPattern.ExceptionCount);
			WORD i = 0;
			for (i = 0; i < arpPattern.ExceptionCount; i++)
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
						Parser.GetStringA(arpPattern.ExceptionInfo[i].SubjectLength2, &arpPattern.ExceptionInfo[i].Subject);
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
						Parser.GetStringA(arpPattern.ExceptionInfo[i].LocationLength2, &arpPattern.ExceptionInfo[i].Location);
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
	Parser.GetBYTES(arpPattern.ReservedBlock1Size, _MaxBytes, &arpPattern.ReservedBlock1);

	if (arpPattern.ExceptionCount &&
		arpPattern.ExceptionCount == arpPattern.RecurrencePattern->ModifiedInstanceCount &&
		arpPattern.ExceptionCount < _MaxEntriesSmall &&
		arpPattern.ExceptionInfo)
	{
		arpPattern.ExtendedException = new ExtendedExceptionStruct[arpPattern.ExceptionCount];
		if (arpPattern.ExtendedException)
		{
			memset(arpPattern.ExtendedException, 0, sizeof(ExtendedExceptionStruct)* arpPattern.ExceptionCount);
			WORD i = 0;
			for (i = 0; i < arpPattern.ExceptionCount; i++)
			{
				if (arpPattern.WriterVersion2 >= 0x0003009)
				{
					Parser.GetDWORD(&arpPattern.ExtendedException[i].ChangeHighlight.ChangeHighlightSize);
					Parser.GetDWORD(&arpPattern.ExtendedException[i].ChangeHighlight.ChangeHighlightValue);
					if (arpPattern.ExtendedException[i].ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
					{
						Parser.GetBYTES(arpPattern.ExtendedException[i].ChangeHighlight.ChangeHighlightSize - sizeof(DWORD), _MaxBytes, &arpPattern.ExtendedException[i].ChangeHighlight.Reserved);
					}
				}
				Parser.GetDWORD(&arpPattern.ExtendedException[i].ReservedBlockEE1Size);
				Parser.GetBYTES(arpPattern.ExtendedException[i].ReservedBlockEE1Size, _MaxBytes, &arpPattern.ExtendedException[i].ReservedBlockEE1);

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
						Parser.GetStringW(arpPattern.ExtendedException[i].WideCharSubjectLength, &arpPattern.ExtendedException[i].WideCharSubject);
					}
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					Parser.GetWORD(&arpPattern.ExtendedException[i].WideCharLocationLength);
					if (arpPattern.ExtendedException[i].WideCharLocationLength)
					{
						Parser.GetStringW(arpPattern.ExtendedException[i].WideCharLocationLength, &arpPattern.ExtendedException[i].WideCharLocation);
					}
				}
				if (arpPattern.ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
					arpPattern.ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					Parser.GetDWORD(&arpPattern.ExtendedException[i].ReservedBlockEE2Size);
					Parser.GetBYTES(arpPattern.ExtendedException[i].ReservedBlockEE2Size, _MaxBytes, &arpPattern.ExtendedException[i].ReservedBlockEE2);

				}
			}
		}
	}
	Parser.GetDWORD(&arpPattern.ReservedBlock2Size);
	Parser.GetBYTES(arpPattern.ReservedBlock2Size, _MaxBytes, &arpPattern.ReservedBlock2);

	arpPattern.JunkDataSize = Parser.GetRemainingData(&arpPattern.JunkData);

	AppointmentRecurrencePatternStruct* parpPattern = new AppointmentRecurrencePatternStruct;
	if (parpPattern)
	{
		*parpPattern = arpPattern;
	}

	return parpPattern;
} // BinToAppointmentRecurrencePatternStruct

void DeleteAppointmentRecurrencePatternStruct(_In_ AppointmentRecurrencePatternStruct* parpPattern)
{
	if (!parpPattern) return;
	DeleteRecurrencePatternStruct(parpPattern->RecurrencePattern);
	int i = 0;
	if (parpPattern->ExceptionCount && parpPattern->ExceptionInfo)
	{
		for (i = 0; i < parpPattern->ExceptionCount; i++)
		{
			delete[] parpPattern->ExceptionInfo[i].Subject;
			delete[] parpPattern->ExceptionInfo[i].Location;
		}
	}
	delete[] parpPattern->ExceptionInfo;
	if (parpPattern->ExceptionCount && parpPattern->ExtendedException)
	{
		for (i = 0; i < parpPattern->ExceptionCount; i++)
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
} // DeleteAppointmentRecurrencePatternStruct

// result allocated with new
// clean up with delete[]
_Check_return_ LPWSTR AppointmentRecurrencePatternStructToString(_In_ AppointmentRecurrencePatternStruct* parpPattern)
{
	if (!parpPattern) return NULL;

	CString szARP;
	CString szTmp;
	LPWSTR szRecurrencePattern = NULL;

	szRecurrencePattern = RecurrencePatternStructToString(parpPattern->RecurrencePattern);
	szARP = szRecurrencePattern;
	delete[] szRecurrencePattern;

	szTmp.FormatMessage(IDS_ARPHEADER,
		parpPattern->ReaderVersion2,
		parpPattern->WriterVersion2,
		parpPattern->StartTimeOffset, RTimeToString(parpPattern->StartTimeOffset),
		parpPattern->EndTimeOffset, RTimeToString(parpPattern->EndTimeOffset),
		parpPattern->ExceptionCount);
	szARP += szTmp;

	WORD i = 0;
	if (parpPattern->ExceptionCount && parpPattern->ExceptionInfo)
	{
		for (i = 0; i < parpPattern->ExceptionCount; i++)
		{
			CString szExceptionInfo;
			LPTSTR szOverrideFlags = NULL;
			InterpretFlags(flagOverrideFlags, parpPattern->ExceptionInfo[i].OverrideFlags, &szOverrideFlags);
			szExceptionInfo.FormatMessage(IDS_ARPEXHEADER,
				i, parpPattern->ExceptionInfo[i].StartDateTime, RTimeToString(parpPattern->ExceptionInfo[i].StartDateTime),
				parpPattern->ExceptionInfo[i].EndDateTime, RTimeToString(parpPattern->ExceptionInfo[i].EndDateTime),
				parpPattern->ExceptionInfo[i].OriginalStartDate, RTimeToString(parpPattern->ExceptionInfo[i].OriginalStartDate),
				parpPattern->ExceptionInfo[i].OverrideFlags, szOverrideFlags);
			delete[] szOverrideFlags;
			szOverrideFlags = NULL;
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
			{
				szTmp.FormatMessage(IDS_ARPEXSUBJECT,
					i, parpPattern->ExceptionInfo[i].SubjectLength,
					parpPattern->ExceptionInfo[i].SubjectLength2,
					parpPattern->ExceptionInfo[i].Subject ? parpPattern->ExceptionInfo[i].Subject : "");
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_MEETINGTYPE)
			{
				LPWSTR szFlags = NULL;
				InterpretNumberAsStringNamedProp(parpPattern->ExceptionInfo[i].MeetingType, dispidApptStateFlags, (LPGUID)&PSETID_Appointment, &szFlags);
				szTmp.FormatMessage(IDS_ARPEXMEETINGTYPE,
					i, parpPattern->ExceptionInfo[i].MeetingType, szFlags);
				delete[] szFlags;
				szFlags = NULL;
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_REMINDERDELTA)
			{
				szTmp.FormatMessage(IDS_ARPEXREMINDERDELTA,
					i, parpPattern->ExceptionInfo[i].ReminderDelta);
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_REMINDER)
			{
				szTmp.FormatMessage(IDS_ARPEXREMINDERSET,
					i, parpPattern->ExceptionInfo[i].ReminderSet);
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				szTmp.FormatMessage(IDS_ARPEXLOCATION,
					i, parpPattern->ExceptionInfo[i].LocationLength,
					parpPattern->ExceptionInfo[i].LocationLength2,
					parpPattern->ExceptionInfo[i].Location ? parpPattern->ExceptionInfo[i].Location : "");
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_BUSYSTATUS)
			{
				LPWSTR szFlags = NULL;
				InterpretNumberAsStringNamedProp(parpPattern->ExceptionInfo[i].BusyStatus, dispidBusyStatus, (LPGUID)&PSETID_Appointment, &szFlags);
				szTmp.FormatMessage(IDS_ARPEXBUSYSTATUS,
					i, parpPattern->ExceptionInfo[i].BusyStatus, szFlags);
				delete[] szFlags;
				szFlags = NULL;
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_ATTACHMENT)
			{
				szTmp.FormatMessage(IDS_ARPEXATTACHMENT,
					i, parpPattern->ExceptionInfo[i].Attachment);
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_SUBTYPE)
			{
				szTmp.FormatMessage(IDS_ARPEXSUBTYPE,
					i, parpPattern->ExceptionInfo[i].SubType);
				szExceptionInfo += szTmp;
			}
			if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_APPTCOLOR)
			{
				szTmp.FormatMessage(IDS_ARPEXAPPOINTMENTCOLOR,
					i, parpPattern->ExceptionInfo[i].AppointmentColor);
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
		SBinary sBin = { 0 };
		sBin.cb = parpPattern->ReservedBlock1Size;
		sBin.lpb = parpPattern->ReservedBlock1;
		szARP += wstringToCString(BinToHexString(&sBin, true));
	}

	if (parpPattern->ExceptionCount && parpPattern->ExtendedException)
	{
		for (i = 0; i < parpPattern->ExceptionCount; i++)
		{
			CString szExtendedException;
			if (parpPattern->WriterVersion2 >= 0x00003009)
			{
				LPWSTR szFlags = NULL;
				InterpretNumberAsStringNamedProp(parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightValue, dispidChangeHighlight, (LPGUID)&PSETID_Appointment, &szFlags);
				szTmp.FormatMessage(IDS_ARPEXCHANGEHIGHLIGHT,
					i, parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightSize,
					parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightValue, szFlags);
				delete[] szFlags;
				szFlags = NULL;
				szExtendedException += szTmp;
				if (parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
				{
					szTmp.FormatMessage(IDS_ARPEXCHANGEHIGHLIGHTRESERVED,
						i);
					szExtendedException += szTmp;
					SBinary sBin = { 0 };
					sBin.cb = parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightSize - sizeof(DWORD);
					sBin.lpb = parpPattern->ExtendedException[i].ChangeHighlight.Reserved;
					szExtendedException += wstringToCString(BinToHexString(&sBin, true));
					szExtendedException += _T("\n"); // STRING_OK
				}
			}
			szTmp.FormatMessage(IDS_ARPEXRESERVED1,
				i, parpPattern->ExtendedException[i].ReservedBlockEE1Size);
			szExtendedException += szTmp;
			if (parpPattern->ExtendedException[i].ReservedBlockEE1Size)
			{
				SBinary sBin = { 0 };
				sBin.cb = parpPattern->ExtendedException[i].ReservedBlockEE1Size;
				sBin.lpb = parpPattern->ExtendedException[i].ReservedBlockEE1;
				szExtendedException += wstringToCString(BinToHexString(&sBin, true));
			}
			if (parpPattern->ExceptionInfo)
			{
				if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
					parpPattern->ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					szTmp.FormatMessage(IDS_ARPEXDATETIME,
						i, parpPattern->ExtendedException[i].StartDateTime, RTimeToString(parpPattern->ExtendedException[i].StartDateTime),
						parpPattern->ExtendedException[i].EndDateTime, RTimeToString(parpPattern->ExtendedException[i].EndDateTime),
						parpPattern->ExtendedException[i].OriginalStartDate, RTimeToString(parpPattern->ExtendedException[i].OriginalStartDate));
					szExtendedException += szTmp;
				}
				if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
				{
					szTmp.FormatMessage(IDS_ARPEXWIDESUBJECT,
						i, parpPattern->ExtendedException[i].WideCharSubjectLength,
						parpPattern->ExtendedException[i].WideCharSubject ? parpPattern->ExtendedException[i].WideCharSubject : L"");
					szExtendedException += szTmp;
				}
				if (parpPattern->ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					szTmp.FormatMessage(IDS_ARPEXWIDELOCATION,
						i, parpPattern->ExtendedException[i].WideCharLocationLength,
						parpPattern->ExtendedException[i].WideCharLocation ? parpPattern->ExtendedException[i].WideCharLocation : L"");
					szExtendedException += szTmp;
				}
			}
			szTmp.FormatMessage(IDS_ARPEXRESERVED1,
				i, parpPattern->ExtendedException[i].ReservedBlockEE2Size);
			szExtendedException += szTmp;
			if (parpPattern->ExtendedException[i].ReservedBlockEE2Size)
			{
				SBinary sBin = { 0 };
				sBin.cb = parpPattern->ExtendedException[i].ReservedBlockEE2Size;
				sBin.lpb = parpPattern->ExtendedException[i].ReservedBlockEE2;
				szExtendedException += wstringToCString(BinToHexString(&sBin, true));
			}

			szARP += szExtendedException;
		}
	}

	szTmp.FormatMessage(IDS_ARPRESERVED2,
		parpPattern->ReservedBlock2Size);
	szARP += szTmp;
	if (parpPattern->ReservedBlock2Size)
	{
		SBinary sBin = { 0 };
		sBin.cb = parpPattern->ReservedBlock2Size;
		sBin.lpb = parpPattern->ReservedBlock2;
		szARP += wstringToCString(BinToHexString(&sBin, true));
	}

	szARP += JunkDataToString(parpPattern->JunkDataSize, parpPattern->JunkData);

	return CStringToLPWSTR(szARP);
} // AppointmentRecurrencePatternStructToString

//////////////////////////////////////////////////////////////////////////
// End AppointmentRecurrencePatternStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// SDBin
//////////////////////////////////////////////////////////////////////////

void SDBinToString(SBinary myBin, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_z_ LPWSTR* lpszResultString)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;

	HRESULT hRes = S_OK;
	LPBYTE lpSDToParse = myBin.lpb;
	ULONG ulSDToParse = myBin.cb;

	if (lpSDToParse)
	{
		eAceType acetype = acetypeMessage;
		switch (GetMAPIObjectType(lpMAPIProp))
		{
		case (MAPI_STORE) :
		case (MAPI_ADDRBOOK) :
		case (MAPI_FOLDER) :
		case (MAPI_ABCONT) :
						   acetype = acetypeContainer;
			break;
		}

		if (PR_FREEBUSY_NT_SECURITY_DESCRIPTOR == ulPropTag)
			acetype = acetypeFreeBusy;

		CString szDACL;
		CString szInfo;
		CString szTmp;

		WC_H(SDToString(lpSDToParse, ulSDToParse, acetype, &szDACL, &szInfo));

		LPTSTR szFlags = NULL;
		InterpretFlags(flagSecurityVersion, SECURITY_DESCRIPTOR_VERSION(lpSDToParse), &szFlags);

		CString szResult;
		szResult.FormatMessage(IDS_SECURITYDESCRIPTORHEADER);
		szResult += szInfo;
		szTmp.FormatMessage(IDS_SECURITYDESCRIPTORVERSION, SECURITY_DESCRIPTOR_VERSION(lpSDToParse), szFlags);
		szResult += szTmp;
		szResult += szDACL;

		delete[] szFlags;
		szFlags = NULL;

		*lpszResultString = CStringToLPWSTR(szResult);
	}
} // SDBinToString

//////////////////////////////////////////////////////////////////////////
// SIDDBin
//////////////////////////////////////////////////////////////////////////

void SIDBinToString(SBinary myBin, _Deref_out_z_ LPWSTR* lpszResultString)
{
	if (!lpszResultString) return;
	HRESULT hRes = S_OK;
	PSID SidStart = myBin.lpb;
	LPTSTR lpSidName = NULL;
	LPTSTR lpSidDomain = NULL;
	LPTSTR lpStringSid = NULL;

	if (SidStart &&
		myBin.cb >= sizeof(SID)-sizeof(DWORD)+sizeof(DWORD)* ((PISID)SidStart)->SubAuthorityCount &&
		IsValidSid(SidStart))
	{
		DWORD dwSidName = 0;
		DWORD dwSidDomain = 0;
		SID_NAME_USE SidNameUse;

		if (!LookupAccountSid(
			NULL,
			SidStart,
			NULL,
			&dwSidName,
			NULL,
			&dwSidDomain,
			&SidNameUse))
		{
			DWORD dwErr = GetLastError();
			hRes = HRESULT_FROM_WIN32(dwErr);
			if (ERROR_NONE_MAPPED != dwErr &&
				STRSAFE_E_INSUFFICIENT_BUFFER != hRes)
			{
				LogFunctionCall(hRes, NULL, false, false, true, dwErr, "LookupAccountSid", __FILE__, __LINE__);
			}
		}
		hRes = S_OK;

#pragma warning(push)
#pragma warning(disable:6211)
		if (dwSidName) lpSidName = new TCHAR[dwSidName];
		if (dwSidDomain) lpSidDomain = new TCHAR[dwSidDomain];
#pragma warning(pop)

		// Only make the call if we got something to get
		if (lpSidName || lpSidDomain)
		{
			WC_B(LookupAccountSid(
				NULL,
				SidStart,
				lpSidName,
				&dwSidName,
				lpSidDomain,
				&dwSidDomain,
				&SidNameUse));
			hRes = S_OK;
		}

		EC_B(GetTextualSid(SidStart, &lpStringSid));
	}

	CString szDomain;
	CString szName;
	CString szSID;

	if (lpSidDomain) szDomain = lpSidDomain;
	else EC_B(szDomain.LoadString(IDS_NODOMAIN));
	if (lpSidName) szName = lpSidName;
	else EC_B(szName.LoadString(IDS_NONAME));
	if (lpStringSid) szSID = lpStringSid;
	else EC_B(szSID.LoadString(IDS_NOSID));

	CString szResult;
	szResult.FormatMessage(IDS_SIDHEADER, szDomain, szName, szSID);

	if (lpStringSid) delete[] lpStringSid;
	if (lpSidDomain) delete[] lpSidDomain;
	if (lpSidName) delete[] lpSidName;

	*lpszResultString = CStringToLPWSTR(szResult);
} // SIDBinToString

//////////////////////////////////////////////////////////////////////////
// ExtendedFlagsStruct
//////////////////////////////////////////////////////////////////////////

_Check_return_ ExtendedFlagsStruct* BinToExtendedFlagsStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	ExtendedFlagsStruct efExtendedFlags = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

	// Run through the parser once to count the number of flag structs
	for (;;)
	{
		// Must have at least 2 bytes left to have another flag
		if (Parser.RemainingBytes() < 2) break;
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
	CBinaryParser Parser2(cbBin, lpBin);
	if (efExtendedFlags.ulNumFlags && efExtendedFlags.ulNumFlags < _MaxEntriesSmall)
		efExtendedFlags.pefExtendedFlags = new ExtendedFlagStruct[efExtendedFlags.ulNumFlags];

	if (efExtendedFlags.pefExtendedFlags)
	{
		memset(efExtendedFlags.pefExtendedFlags, 0, sizeof(ExtendedFlagStruct)*efExtendedFlags.ulNumFlags);
		ULONG i = 0;
		bool bBadData = false;

		for (i = 0; i < efExtendedFlags.ulNumFlags; i++)
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
					Parser2.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)&efExtendedFlags.pefExtendedFlags[i].Data.SearchFolderID);
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
				Parser2.GetBYTES(efExtendedFlags.pefExtendedFlags[i].Cb, _MaxBytes, &efExtendedFlags.pefExtendedFlags[i].lpUnknownData);
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

	efExtendedFlags.JunkDataSize = Parser2.GetRemainingData(&efExtendedFlags.JunkData);

	ExtendedFlagsStruct* pefExtendedFlags = new ExtendedFlagsStruct;
	if (pefExtendedFlags)
	{
		*pefExtendedFlags = efExtendedFlags;
	}

	return pefExtendedFlags;
} // BinToExtendedFlagsStruct

void DeleteExtendedFlagsStruct(_In_ ExtendedFlagsStruct* pefExtendedFlags)
{
	if (!pefExtendedFlags) return;
	ULONG i = 0;
	if (pefExtendedFlags->ulNumFlags && pefExtendedFlags->pefExtendedFlags)
	{
		for (i = 0; i < pefExtendedFlags->ulNumFlags; i++)
		{
			delete[] pefExtendedFlags->pefExtendedFlags[i].lpUnknownData;
		}
	}
	delete[] pefExtendedFlags->pefExtendedFlags;
	delete[] pefExtendedFlags->JunkData;
	delete pefExtendedFlags;
} // DeleteExtendedFlagsStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR ExtendedFlagsStructToString(_In_ ExtendedFlagsStruct* pefExtendedFlags)
{
	if (!pefExtendedFlags) return NULL;

	CString szExtendedFlags;
	CString szTmp;
	SBinary sBin = { 0 };

	szExtendedFlags.FormatMessage(IDS_EXTENDEDFLAGSHEADER, pefExtendedFlags->ulNumFlags);

	if (pefExtendedFlags->ulNumFlags && pefExtendedFlags->pefExtendedFlags)
	{
		ULONG i = 0;
		for (i = 0; i < pefExtendedFlags->ulNumFlags; i++)
		{
			LPTSTR szFlags = NULL;
			InterpretFlags(flagExtendedFolderFlagType, pefExtendedFlags->pefExtendedFlags[i].Id, &szFlags);
			szTmp.FormatMessage(IDS_EXTENDEDFLAGID,
				pefExtendedFlags->pefExtendedFlags[i].Id, szFlags,
				pefExtendedFlags->pefExtendedFlags[i].Cb);
			delete[] szFlags;
			szFlags = NULL;
			szExtendedFlags += szTmp;

			switch (pefExtendedFlags->pefExtendedFlags[i].Id)
			{
			case EFPB_FLAGS:
				InterpretFlags(flagExtendedFolderFlag, pefExtendedFlags->pefExtendedFlags[i].Data.ExtendedFlags, &szFlags);
				szTmp.FormatMessage(IDS_EXTENDEDFLAGDATAFLAG, pefExtendedFlags->pefExtendedFlags[i].Data.ExtendedFlags, szFlags);
				delete[] szFlags;
				szFlags = NULL;
				szExtendedFlags += szTmp;
				break;
			case EFPB_CLSIDID:
				szFlags = GUIDToString(&pefExtendedFlags->pefExtendedFlags[i].Data.SearchFolderID);
				szTmp.FormatMessage(IDS_EXTENDEDFLAGDATASFID, szFlags);
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
				szTmp.FormatMessage(IDS_EXTENDEDFLAGDATATODOVERSION, pefExtendedFlags->pefExtendedFlags[i].Data.ToDoFolderVersion);
				szExtendedFlags += szTmp;
				break;
			}
			if (pefExtendedFlags->pefExtendedFlags[i].lpUnknownData)
			{
				szTmp.FormatMessage(IDS_EXTENDEDFLAGUNKNOWN);
				szExtendedFlags += szTmp;
				sBin.cb = pefExtendedFlags->pefExtendedFlags[i].Cb;
				sBin.lpb = pefExtendedFlags->pefExtendedFlags[i].lpUnknownData;
				szExtendedFlags += wstringToCString(BinToHexString(&sBin, true));
			}
		}
	}

	szExtendedFlags += JunkDataToString(pefExtendedFlags->JunkDataSize, pefExtendedFlags->JunkData);

	return CStringToLPWSTR(szExtendedFlags);
} // ExtendedFlagsStructToString

//////////////////////////////////////////////////////////////////////////
// End ExtendedFlagsStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// TimeZoneStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new. Clean up with DeleteTimeZoneStruct.
_Check_return_ TimeZoneStruct* BinToTimeZoneStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	TimeZoneStruct tzTimeZone = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

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

	tzTimeZone.JunkDataSize = Parser.GetRemainingData(&tzTimeZone.JunkData);

	TimeZoneStruct* ptzTimeZone = new TimeZoneStruct;
	if (ptzTimeZone)
	{
		*ptzTimeZone = tzTimeZone;
	}

	return ptzTimeZone;
} // BinToTimeZoneStruct

void DeleteTimeZoneStruct(_In_ TimeZoneStruct* ptzTimeZone)
{
	if (!ptzTimeZone) return;
	delete[] ptzTimeZone->JunkData;
	delete ptzTimeZone;
} // DeleteTimeZoneStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR TimeZoneStructToString(_In_ TimeZoneStruct* ptzTimeZone)
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

	szTimeZone += JunkDataToString(ptzTimeZone->JunkDataSize, ptzTimeZone->JunkData);

	return CStringToLPWSTR(szTimeZone);
} // TimeZoneStructToString

//////////////////////////////////////////////////////////////////////////
// End TimeZoneStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// TimeZoneDefinitionStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new. Clean up with DeleteTimeZoneDefinitionStruct.
_Check_return_ TimeZoneDefinitionStruct* BinToTimeZoneDefinitionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	TimeZoneDefinitionStruct tzdTimeZoneDefinition = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

	Parser.GetBYTE(&tzdTimeZoneDefinition.bMajorVersion);
	Parser.GetBYTE(&tzdTimeZoneDefinition.bMinorVersion);
	Parser.GetWORD(&tzdTimeZoneDefinition.cbHeader);
	Parser.GetWORD(&tzdTimeZoneDefinition.wReserved);
	Parser.GetWORD(&tzdTimeZoneDefinition.cchKeyName);
	Parser.GetStringW(tzdTimeZoneDefinition.cchKeyName, &tzdTimeZoneDefinition.szKeyName);
	Parser.GetWORD(&tzdTimeZoneDefinition.cRules);

	if (tzdTimeZoneDefinition.cRules && tzdTimeZoneDefinition.cRules < _MaxEntriesSmall)
		tzdTimeZoneDefinition.lpTZRule = new TZRule[tzdTimeZoneDefinition.cRules];

	if (tzdTimeZoneDefinition.lpTZRule)
	{
		memset(tzdTimeZoneDefinition.lpTZRule, 0, sizeof(TZRule)* tzdTimeZoneDefinition.cRules);
		ULONG i = 0;
		for (i = 0; i < tzdTimeZoneDefinition.cRules; i++)
		{
			Parser.GetBYTE(&tzdTimeZoneDefinition.lpTZRule[i].bMajorVersion);
			Parser.GetBYTE(&tzdTimeZoneDefinition.lpTZRule[i].bMinorVersion);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].wReserved);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].wTZRuleFlags);
			Parser.GetWORD(&tzdTimeZoneDefinition.lpTZRule[i].wYear);
			Parser.GetBYTESNoAlloc(sizeof(tzdTimeZoneDefinition.lpTZRule[i].X), sizeof(tzdTimeZoneDefinition.lpTZRule[i].X), tzdTimeZoneDefinition.lpTZRule[i].X);
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

	tzdTimeZoneDefinition.JunkDataSize = Parser.GetRemainingData(&tzdTimeZoneDefinition.JunkData);

	TimeZoneDefinitionStruct* ptzdTimeZoneDefinition = new TimeZoneDefinitionStruct;
	if (ptzdTimeZoneDefinition)
	{
		*ptzdTimeZoneDefinition = tzdTimeZoneDefinition;
	}

	return ptzdTimeZoneDefinition;
} // BinToTimeZoneDefinitionStruct

void DeleteTimeZoneDefinitionStruct(_In_ TimeZoneDefinitionStruct* ptzdTimeZoneDefinition)
{
	if (!ptzdTimeZoneDefinition) return;
	delete[] ptzdTimeZoneDefinition->szKeyName;
	delete[] ptzdTimeZoneDefinition->lpTZRule;
	delete[] ptzdTimeZoneDefinition->JunkData;
	delete ptzdTimeZoneDefinition;
} // DeleteTimeZoneDefinitionStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR TimeZoneDefinitionStructToString(_In_ TimeZoneDefinitionStruct* ptzdTimeZoneDefinition)
{
	if (!ptzdTimeZoneDefinition) return NULL;

	CString szTimeZoneDefinition;
	CString szTmp;

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

		for (i = 0; i < ptzdTimeZoneDefinition->cRules; i++)
		{
			LPTSTR szFlags = NULL;
			InterpretFlags(flagTZRule, ptzdTimeZoneDefinition->lpTZRule[i].wTZRuleFlags, &szFlags);
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

			SBinary sBin = { 0 };
			sBin.cb = 14;
			sBin.lpb = ptzdTimeZoneDefinition->lpTZRule[i].X;
			szTZRule += wstringToCString(BinToHexString(&sBin, true));

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

	szTimeZoneDefinition += JunkDataToString(ptzdTimeZoneDefinition->JunkDataSize, ptzdTimeZoneDefinition->JunkData);

	return CStringToLPWSTR(szTimeZoneDefinition);
} // TimeZoneDefinitionStructToString

//////////////////////////////////////////////////////////////////////////
// End TimeZoneDefinitionStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// ReportTagStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new. Clean up with DeleteReportTagStruct.
_Check_return_ ReportTagStruct* BinToReportTagStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	ReportTagStruct rtReportTag = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

	Parser.GetBYTESNoAlloc(sizeof(rtReportTag.Cookie), sizeof(rtReportTag.Cookie), (LPBYTE)rtReportTag.Cookie);

	// Version is big endian, so we have to read individual bytes
	WORD hiWord = NULL;
	WORD loWord = NULL;
	Parser.GetWORD(&hiWord);
	Parser.GetWORD(&loWord);
	rtReportTag.Version = (hiWord << 16) | loWord;

	Parser.GetDWORD(&rtReportTag.cbStoreEntryID);
	if (rtReportTag.cbStoreEntryID)
	{
		Parser.GetBYTES(rtReportTag.cbStoreEntryID, _MaxEID, &rtReportTag.lpStoreEntryID);
	}
	Parser.GetDWORD(&rtReportTag.cbFolderEntryID);
	if (rtReportTag.cbFolderEntryID)
	{
		Parser.GetBYTES(rtReportTag.cbFolderEntryID, _MaxEID, &rtReportTag.lpFolderEntryID);
	}
	Parser.GetDWORD(&rtReportTag.cbMessageEntryID);
	if (rtReportTag.cbMessageEntryID)
	{
		Parser.GetBYTES(rtReportTag.cbMessageEntryID, _MaxEID, &rtReportTag.lpMessageEntryID);
	}
	Parser.GetDWORD(&rtReportTag.cbSearchFolderEntryID);
	if (rtReportTag.cbSearchFolderEntryID)
	{
		Parser.GetBYTES(rtReportTag.cbSearchFolderEntryID, _MaxEID, &rtReportTag.lpSearchFolderEntryID);
	}
	Parser.GetDWORD(&rtReportTag.cbMessageSearchKey);
	if (rtReportTag.cbMessageSearchKey)
	{
		Parser.GetBYTES(rtReportTag.cbMessageSearchKey, _MaxEID, &rtReportTag.lpMessageSearchKey);
	}
	Parser.GetDWORD(&rtReportTag.cchAnsiText);
	if (rtReportTag.cchAnsiText)
	{
		Parser.GetStringA(rtReportTag.cchAnsiText, &rtReportTag.lpszAnsiText);
	}

	rtReportTag.JunkDataSize = Parser.GetRemainingData(&rtReportTag.JunkData);

	ReportTagStruct* prtReportTag = new ReportTagStruct;
	if (prtReportTag)
	{
		*prtReportTag = rtReportTag;
	}

	return prtReportTag;
} // BinToReportTagStruct

void DeleteReportTagStruct(_In_ ReportTagStruct* prtReportTag)
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
} // DeleteReportTagStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR ReportTagStructToString(_In_ ReportTagStruct* prtReportTag)
{
	if (!prtReportTag) return NULL;

	CString szReportTag;
	CString szTmp;

	szReportTag.FormatMessage(IDS_REPORTTAGHEADER);

	SBinary sBin = { 0 };
	sBin.cb = sizeof(prtReportTag->Cookie);
	sBin.lpb = (LPBYTE)prtReportTag->Cookie;
	szReportTag += wstringToCString(BinToHexString(&sBin, true));

	LPTSTR szFlags = NULL;
	InterpretFlags(flagReportTagVersion, prtReportTag->Version, &szFlags);
	szTmp.FormatMessage(IDS_REPORTTAGVERSION,
		prtReportTag->Version,
		szFlags);
	delete[] szFlags;
	szFlags = NULL;
	szReportTag += szTmp;

	if (prtReportTag->cbStoreEntryID)
	{
		szTmp.FormatMessage(IDS_REPORTTAGSTOREEID);
		szReportTag += szTmp;
		sBin.cb = prtReportTag->cbStoreEntryID;
		sBin.lpb = prtReportTag->lpStoreEntryID;
		szReportTag += wstringToCString(BinToHexString(&sBin, true));
	}

	if (prtReportTag->cbFolderEntryID)
	{
		szTmp.FormatMessage(IDS_REPORTTAGFOLDEREID);
		szReportTag += szTmp;
		sBin.cb = prtReportTag->cbFolderEntryID;
		sBin.lpb = prtReportTag->lpFolderEntryID;
		szReportTag += wstringToCString(BinToHexString(&sBin, true));
	}

	if (prtReportTag->cbMessageEntryID)
	{
		szTmp.FormatMessage(IDS_REPORTTAGMESSAGEEID);
		szReportTag += szTmp;
		sBin.cb = prtReportTag->cbMessageEntryID;
		sBin.lpb = prtReportTag->lpMessageEntryID;
		szReportTag += wstringToCString(BinToHexString(&sBin, true));
	}

	if (prtReportTag->cbSearchFolderEntryID)
	{
		szTmp.FormatMessage(IDS_REPORTTAGSFEID);
		szReportTag += szTmp;
		sBin.cb = prtReportTag->cbSearchFolderEntryID;
		sBin.lpb = prtReportTag->lpSearchFolderEntryID;
		szReportTag += wstringToCString(BinToHexString(&sBin, true));
	}

	if (prtReportTag->cbMessageSearchKey)
	{
		szTmp.FormatMessage(IDS_REPORTTAGMESSAGEKEY);
		szReportTag += szTmp;
		sBin.cb = prtReportTag->cbMessageSearchKey;
		sBin.lpb = prtReportTag->lpMessageSearchKey;
		szReportTag += wstringToCString(BinToHexString(&sBin, true));
	}

	if (prtReportTag->cchAnsiText)
	{
		szTmp.FormatMessage(IDS_REPORTTAGANSITEXT,
			prtReportTag->cchAnsiText,
			prtReportTag->lpszAnsiText ? prtReportTag->lpszAnsiText : ""); // STRING_OK
		szReportTag += szTmp;
	}

	szReportTag += JunkDataToString(prtReportTag->JunkDataSize, prtReportTag->JunkData);

	return CStringToLPWSTR(szReportTag);
} // ReportTagStructToString

//////////////////////////////////////////////////////////////////////////
// End ReportTagStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// ConversationIndexStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new. Clean up with DeleteConversationIndexStruct.
_Check_return_ ConversationIndexStruct* BinToConversationIndexStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	ConversationIndexStruct ciConversationIndex = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

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
	ciConversationIndex.guid.Data2 = (unsigned short)((b1 << 8) | b2);
	Parser.GetBYTE(&b1);
	Parser.GetBYTE(&b2);
	ciConversationIndex.guid.Data3 = (unsigned short)((b1 << 8) | b2);
	Parser.GetBYTESNoAlloc(sizeof(ciConversationIndex.guid.Data4), sizeof(ciConversationIndex.guid.Data4), ciConversationIndex.guid.Data4);

	if (Parser.RemainingBytes() > 0)
	{
		ciConversationIndex.ulResponseLevels = (ULONG)Parser.RemainingBytes() / 5; // Response levels consume 5 bytes each
	}

	if (ciConversationIndex.ulResponseLevels && ciConversationIndex.ulResponseLevels < _MaxEntriesSmall)
		ciConversationIndex.lpResponseLevels = new ResponseLevelStruct[ciConversationIndex.ulResponseLevels];

	if (ciConversationIndex.lpResponseLevels)
	{
		memset(ciConversationIndex.lpResponseLevels, 0, sizeof(ResponseLevelStruct)* ciConversationIndex.ulResponseLevels);
		ULONG i = 0;
		for (i = 0; i < ciConversationIndex.ulResponseLevels; i++)
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
			ciConversationIndex.lpResponseLevels[i].Random = (BYTE)(b1 >> 4);
			ciConversationIndex.lpResponseLevels[i].ResponseLevel = (BYTE)(b1 & 0xf);
		}
	}

	ciConversationIndex.JunkDataSize = Parser.GetRemainingData(&ciConversationIndex.JunkData);

	ConversationIndexStruct* pciConversationIndex = new ConversationIndexStruct;
	if (pciConversationIndex)
	{
		*pciConversationIndex = ciConversationIndex;
	}

	return pciConversationIndex;
} // BinToConversationIndexStruct

void DeleteConversationIndexStruct(_In_ ConversationIndexStruct* pciConversationIndex)
{
	if (!pciConversationIndex) return;
	delete[] pciConversationIndex->lpResponseLevels;
	delete[] pciConversationIndex->JunkData;
	delete pciConversationIndex;
} // DeleteConversationIndexStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR ConversationIndexStructToString(_In_ ConversationIndexStruct* pciConversationIndex)
{
	if (!pciConversationIndex) return NULL;

	CString szConversationIndex;
	CString szTmp;

	CString PropString;
	FileTimeToString(&pciConversationIndex->ftCurrent, &PropString, NULL);
	LPTSTR szGUID = GUIDToString(&pciConversationIndex->guid);
	szTmp.FormatMessage(IDS_CONVERSATIONINDEXHEADER,
		pciConversationIndex->UnnamedByte,
		pciConversationIndex->ftCurrent.dwLowDateTime,
		pciConversationIndex->ftCurrent.dwHighDateTime,
		PropString,
		szGUID);
	szConversationIndex += szTmp;
	delete[] szGUID;

	if (pciConversationIndex->ulResponseLevels && pciConversationIndex->lpResponseLevels)
	{
		ULONG i = 0;
		for (i = 0; i < pciConversationIndex->ulResponseLevels; i++)
		{
			szTmp.FormatMessage(IDS_CONVERSATIONINDEXRESPONSELEVEL,
				i, pciConversationIndex->lpResponseLevels[i].DeltaCode,
				pciConversationIndex->lpResponseLevels[i].TimeDelta,
				pciConversationIndex->lpResponseLevels[i].Random,
				pciConversationIndex->lpResponseLevels[i].ResponseLevel);
			szConversationIndex += szTmp;
		}
	}

	szConversationIndex += JunkDataToString(pciConversationIndex->JunkDataSize, pciConversationIndex->JunkData);

	return CStringToLPWSTR(szConversationIndex);
} // ConversationIndexStructToString

//////////////////////////////////////////////////////////////////////////
// ConversationIndexStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// TaskAssignersStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new. Clean up with DeleteTaskAssignersStruct.
_Check_return_ TaskAssignersStruct* BinToTaskAssignersStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	TaskAssignersStruct taTaskAssigners = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

	Parser.GetDWORD(&taTaskAssigners.cAssigners);

	if (taTaskAssigners.cAssigners && taTaskAssigners.cAssigners < _MaxEntriesSmall)
		taTaskAssigners.lpTaskAssigners = new TaskAssignerStruct[taTaskAssigners.cAssigners];

	if (taTaskAssigners.lpTaskAssigners)
	{
		memset(taTaskAssigners.lpTaskAssigners, 0, sizeof(TaskAssignerStruct)* taTaskAssigners.cAssigners);
		DWORD i = 0;
		for (i = 0; i < taTaskAssigners.cAssigners; i++)
		{
			Parser.GetDWORD(&taTaskAssigners.lpTaskAssigners[i].cbAssigner);
			LPBYTE lpbAssigner = NULL;
			Parser.GetBYTES(taTaskAssigners.lpTaskAssigners[i].cbAssigner, _MaxEntriesSmall, &lpbAssigner);
			if (lpbAssigner)
			{
				CBinaryParser AssignerParser(taTaskAssigners.lpTaskAssigners[i].cbAssigner, lpbAssigner);
				AssignerParser.GetDWORD(&taTaskAssigners.lpTaskAssigners[i].cbEntryID);
				AssignerParser.GetBYTES(taTaskAssigners.lpTaskAssigners[i].cbEntryID, _MaxEID, &taTaskAssigners.lpTaskAssigners[i].lpEntryID);
				AssignerParser.GetStringA(&taTaskAssigners.lpTaskAssigners[i].szDisplayName);
				AssignerParser.GetStringW(&taTaskAssigners.lpTaskAssigners[i].wzDisplayName);

				taTaskAssigners.JunkDataSize = AssignerParser.GetRemainingData(&taTaskAssigners.JunkData);
				delete[] lpbAssigner;
			}
		}
	}

	taTaskAssigners.JunkDataSize = Parser.GetRemainingData(&taTaskAssigners.JunkData);

	TaskAssignersStruct* ptaTaskAssigners = new TaskAssignersStruct;
	if (ptaTaskAssigners)
	{
		*ptaTaskAssigners = taTaskAssigners;
	}

	return ptaTaskAssigners;
} // BinToTaskAssignersStruct

void DeleteTaskAssignersStruct(_In_ TaskAssignersStruct* ptaTaskAssigners)
{
	if (!ptaTaskAssigners) return;
	DWORD i = 0;
	if (ptaTaskAssigners->cAssigners && ptaTaskAssigners->lpTaskAssigners)
	{
		for (i = 0; i < ptaTaskAssigners->cAssigners; i++)
		{
			delete[] ptaTaskAssigners->lpTaskAssigners[i].lpEntryID;
			delete[] ptaTaskAssigners->lpTaskAssigners[i].szDisplayName;
			delete[] ptaTaskAssigners->lpTaskAssigners[i].wzDisplayName;
			delete[] ptaTaskAssigners->lpTaskAssigners[i].JunkData;
		}
	}
	delete[] ptaTaskAssigners->lpTaskAssigners;
	delete[] ptaTaskAssigners->JunkData;
	delete ptaTaskAssigners;
} // DeleteTaskAssignersStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR TaskAssignersStructToString(_In_ TaskAssignersStruct* ptaTaskAssigners)
{
	if (!ptaTaskAssigners) return NULL;

	CString szTaskAssigners;
	CString szTmp;

	szTaskAssigners.FormatMessage(IDS_TASKASSIGNERSHEADER,
		ptaTaskAssigners->cAssigners);

	if (ptaTaskAssigners->cAssigners && ptaTaskAssigners->lpTaskAssigners)
	{
		DWORD i = 0;
		for (i = 0; i < ptaTaskAssigners->cAssigners; i++)
		{
			szTmp.FormatMessage(IDS_TASKASSIGNEREID,
				i,
				ptaTaskAssigners->lpTaskAssigners[i].cbEntryID);
			szTaskAssigners += szTmp;
			if (ptaTaskAssigners->lpTaskAssigners[i].lpEntryID)
			{
				SBinary sBin = { 0 };
				sBin.cb = ptaTaskAssigners->lpTaskAssigners[i].cbEntryID;
				sBin.lpb = ptaTaskAssigners->lpTaskAssigners[i].lpEntryID;
				szTaskAssigners += wstringToCString(BinToHexString(&sBin, true));
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
				SBinary sBin = { 0 };
				sBin.cb = (ULONG)ptaTaskAssigners->lpTaskAssigners[i].JunkDataSize;
				sBin.lpb = ptaTaskAssigners->lpTaskAssigners[i].JunkData;
				szTaskAssigners += wstringToCString(BinToHexString(&sBin, true));
			}
		}
	}

	szTaskAssigners += JunkDataToString(ptaTaskAssigners->JunkDataSize, ptaTaskAssigners->JunkData);

	return CStringToLPWSTR(szTaskAssigners);
} // TaskAssignersStructToString

//////////////////////////////////////////////////////////////////////////
// End TaskAssignersStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// GlobalObjectIdStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new. Clean up with DeleteGlobalObjectIdStruct.
_Check_return_ GlobalObjectIdStruct* BinToGlobalObjectIdStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	GlobalObjectIdStruct goidGlobalObjectId = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

	Parser.GetBYTESNoAlloc(sizeof(goidGlobalObjectId.Id), sizeof(goidGlobalObjectId.Id), (LPBYTE)&goidGlobalObjectId.Id);
	BYTE b1 = NULL;
	BYTE b2 = NULL;
	Parser.GetBYTE(&b1);
	Parser.GetBYTE(&b2);
	goidGlobalObjectId.Year = (WORD)((b1 << 8) | b2);
	Parser.GetBYTE(&goidGlobalObjectId.Month);
	Parser.GetBYTE(&goidGlobalObjectId.Day);
	Parser.GetLARGE_INTEGER((LARGE_INTEGER*)&goidGlobalObjectId.CreationTime);
	Parser.GetLARGE_INTEGER(&goidGlobalObjectId.X);
	Parser.GetDWORD(&goidGlobalObjectId.dwSize);
	Parser.GetBYTES(goidGlobalObjectId.dwSize, _MaxBytes, &goidGlobalObjectId.lpData);

	goidGlobalObjectId.JunkDataSize = Parser.GetRemainingData(&goidGlobalObjectId.JunkData);

	GlobalObjectIdStruct* pgoidGlobalObjectId = new GlobalObjectIdStruct;
	if (pgoidGlobalObjectId)
	{
		*pgoidGlobalObjectId = goidGlobalObjectId;
	}

	return pgoidGlobalObjectId;
} // BinToGlobalObjectIdStruct

void DeleteGlobalObjectIdStruct(_In_opt_ GlobalObjectIdStruct* pgoidGlobalObjectId)
{
	if (!pgoidGlobalObjectId) return;
	delete[] pgoidGlobalObjectId->lpData;
	delete[] pgoidGlobalObjectId->JunkData;
	delete pgoidGlobalObjectId;
} // DeleteGlobalObjectIdStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR GlobalObjectIdStructToString(_In_ GlobalObjectIdStruct* pgoidGlobalObjectId)
{
	if (!pgoidGlobalObjectId) return NULL;

	CString szGlobalObjectId;
	CString szTmp;

	szGlobalObjectId.FormatMessage(IDS_GLOBALOBJECTIDHEADER);

	SBinary sBin = { 0 };
	sBin.cb = sizeof(pgoidGlobalObjectId->Id);
	sBin.lpb = pgoidGlobalObjectId->Id;
	szGlobalObjectId += wstringToCString(BinToHexString(&sBin, true));

	LPTSTR szFlags = NULL;
	InterpretFlags(flagGlobalObjectIdMonth, pgoidGlobalObjectId->Month, &szFlags);

	CString PropString;
	FileTimeToString(&pgoidGlobalObjectId->CreationTime, &PropString, NULL);
	szTmp.FormatMessage(IDS_GLOBALOBJECTIDDATA1,
		pgoidGlobalObjectId->Year,
		pgoidGlobalObjectId->Month, szFlags,
		pgoidGlobalObjectId->Day,
		pgoidGlobalObjectId->CreationTime.dwHighDateTime, pgoidGlobalObjectId->CreationTime.dwLowDateTime, PropString,
		pgoidGlobalObjectId->X.HighPart, pgoidGlobalObjectId->X.LowPart,
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
		szGlobalObjectId += wstringToCString(BinToHexString(&sBin, true));
	}

	szGlobalObjectId += JunkDataToString(pgoidGlobalObjectId->JunkDataSize, pgoidGlobalObjectId->JunkData);

	return CStringToLPWSTR(szGlobalObjectId);
} // GlobalObjectIdStructToString

//////////////////////////////////////////////////////////////////////////
// End GlobalObjectIdStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// BinToEntryIdStruct
//////////////////////////////////////////////////////////////////////////

// Fills out lpEID based on the passed in entry ID type
// lpcbBytesRead returns the number of bytes consumed
void BinToTypedEntryIdStruct(EIDStructType ulType, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _In_ EntryIdStruct* lpEID, _Out_ size_t* lpcbBytesRead)
{
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	if (!lpBin || !lpEID) return;

	CBinaryParser Parser(cbBin, lpBin);

	switch (ulType)
	{
		// One Off Recipient
	case eidtOneOff:
	{
					   Parser.GetDWORD(&lpEID->ProviderData.OneOffRecipientObject.Bitmask);
					   if (MAPI_UNICODE & lpEID->ProviderData.OneOffRecipientObject.Bitmask)
					   {
						   Parser.GetStringW(&lpEID->ProviderData.OneOffRecipientObject.Strings.Unicode.DisplayName);
						   Parser.GetStringW(&lpEID->ProviderData.OneOffRecipientObject.Strings.Unicode.AddressType);
						   Parser.GetStringW(&lpEID->ProviderData.OneOffRecipientObject.Strings.Unicode.EmailAddress);
					   }
					   else
					   {
						   Parser.GetStringA(&lpEID->ProviderData.OneOffRecipientObject.Strings.ANSI.DisplayName);
						   Parser.GetStringA(&lpEID->ProviderData.OneOffRecipientObject.Strings.ANSI.AddressType);
						   Parser.GetStringA(&lpEID->ProviderData.OneOffRecipientObject.Strings.ANSI.EmailAddress);
					   }
	}
		break;
		// Address Book Recipient
	case eidtAddressBook:
	{
							Parser.GetDWORD(&lpEID->ProviderData.AddressBookObject.Version);
							Parser.GetDWORD(&lpEID->ProviderData.AddressBookObject.Type);
							Parser.GetStringA(&lpEID->ProviderData.AddressBookObject.X500DN);
	}
		break;
		// Contact Address Book / Personal Distribution List (PDL)
	case eidtContact:
	{
						Parser.GetDWORD(&lpEID->ProviderData.ContactAddressBookObject.Version);
						Parser.GetDWORD(&lpEID->ProviderData.ContactAddressBookObject.Type);

						if (CONTAB_CONTAINER == lpEID->ProviderData.ContactAddressBookObject.Type)
						{
							Parser.GetBYTESNoAlloc(
								sizeof(lpEID->ProviderData.ContactAddressBookObject.muidID),
								sizeof(lpEID->ProviderData.ContactAddressBookObject.muidID),
								(LPBYTE)&lpEID->ProviderData.ContactAddressBookObject.muidID);
						}
						else // Assume we've got some variation on the user/distlist format
						{
							Parser.GetDWORD(&lpEID->ProviderData.ContactAddressBookObject.Index);
							Parser.GetDWORD(&lpEID->ProviderData.ContactAddressBookObject.EntryIDCount);
						}

						// Read the wrapped entry ID from the remaining data
						size_t cbOffset = Parser.GetCurrentOffset();
						size_t cbRemainingBytes = Parser.RemainingBytes();

						// If we already got a size, use it, else we just read the rest of the structure
						if (0 != lpEID->ProviderData.ContactAddressBookObject.EntryIDCount &&
							lpEID->ProviderData.ContactAddressBookObject.EntryIDCount < cbRemainingBytes)
						{
							cbRemainingBytes = lpEID->ProviderData.ContactAddressBookObject.EntryIDCount;
						}
						lpEID->ProviderData.ContactAddressBookObject.lpEntryID = BinToEntryIdStruct(
							(ULONG)cbRemainingBytes,
							lpBin + cbOffset);
						Parser.Advance(cbRemainingBytes);
	}
		break;
	case eidtWAB:
	{
					lpEID->ObjectType = eidtWAB;

					Parser.GetBYTE(&lpEID->ProviderData.WAB.Type);

					size_t cbBinRead = 0;
					lpEID->ProviderData.WAB.lpEntryID = BinToEntryIdStructWithSize(
						(ULONG)Parser.RemainingBytes(),
						lpBin + Parser.GetCurrentOffset(),
						&cbBinRead);
					Parser.Advance(cbBinRead);
	}
		break;
		// message store objects
	case eidtMessageDatabase:
	{
								Parser.GetBYTE(&lpEID->ProviderData.MessageDatabaseObject.Version);
								Parser.GetBYTE(&lpEID->ProviderData.MessageDatabaseObject.Flag);
								Parser.GetStringA(&lpEID->ProviderData.MessageDatabaseObject.DLLFileName);
								lpEID->ProviderData.MessageDatabaseObject.bIsExchange = false;

								// We only know how to parse emsmdb.dll's wrapped entry IDs
								if (lpEID->ProviderData.MessageDatabaseObject.DLLFileName &&
									CSTR_EQUAL == CompareStringA(LOCALE_INVARIANT,
									NORM_IGNORECASE,
									lpEID->ProviderData.MessageDatabaseObject.DLLFileName,
									-1,
									"emsmdb.dll", // STRING_OK
									-1))
								{
									lpEID->ProviderData.MessageDatabaseObject.bIsExchange = true;
									size_t cbRead = Parser.GetCurrentOffset();
									// Advance to the next multiple of 4
									Parser.Advance(3 - ((cbRead + 3) % 4));
									Parser.GetDWORD(&lpEID->ProviderData.MessageDatabaseObject.WrappedFlags);
									Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.MessageDatabaseObject.WrappedProviderUID),
										sizeof(lpEID->ProviderData.MessageDatabaseObject.WrappedProviderUID),
										(LPBYTE)&lpEID->ProviderData.MessageDatabaseObject.WrappedProviderUID);
									Parser.GetDWORD(&lpEID->ProviderData.MessageDatabaseObject.WrappedType);
									Parser.GetStringA(&lpEID->ProviderData.MessageDatabaseObject.ServerShortname);

									lpEID->ProviderData.MessageDatabaseObject.bV2 = false;

									// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
									if (lpEID->ProviderData.MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC)
									{
										cbRead = Parser.GetCurrentOffset();
										Parser.GetDWORD(&lpEID->ProviderData.MessageDatabaseObject.v2.ulMagic);
										if (MDB_STORE_EID_V2_MAGIC == lpEID->ProviderData.MessageDatabaseObject.v2.ulMagic)
										{
											lpEID->ProviderData.MessageDatabaseObject.bV2 = true;
										}
										Parser.SetCurrentOffset(cbRead);
									}

									// Either we're not a PF eid, or this PF EID wasn't followed directly by a magic value
									if (!(lpEID->ProviderData.MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC) ||
										!lpEID->ProviderData.MessageDatabaseObject.bV2)
									{
										Parser.GetStringA(&lpEID->ProviderData.MessageDatabaseObject.MailboxDN);
									}
									if (Parser.RemainingBytes() >= sizeof(MDB_STORE_EID_V2)+sizeof(WCHAR))
									{
										lpEID->ProviderData.MessageDatabaseObject.bV2 = true;
										Parser.GetDWORD(&lpEID->ProviderData.MessageDatabaseObject.v2.ulMagic);
										Parser.GetDWORD(&lpEID->ProviderData.MessageDatabaseObject.v2.ulSize);
										Parser.GetDWORD(&lpEID->ProviderData.MessageDatabaseObject.v2.ulVersion);
										Parser.GetDWORD(&lpEID->ProviderData.MessageDatabaseObject.v2.ulOffsetDN);
										Parser.GetDWORD(&lpEID->ProviderData.MessageDatabaseObject.v2.ulOffsetFQDN);
										if (lpEID->ProviderData.MessageDatabaseObject.v2.ulOffsetDN)
										{
											Parser.GetStringA(&lpEID->ProviderData.MessageDatabaseObject.v2DN);
										}
										if (lpEID->ProviderData.MessageDatabaseObject.v2.ulOffsetFQDN)
										{
											Parser.GetStringW(&lpEID->ProviderData.MessageDatabaseObject.v2FQDN);
										}
										Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.MessageDatabaseObject.v2Reserved),
											sizeof(lpEID->ProviderData.MessageDatabaseObject.v2Reserved),
											(LPBYTE)&lpEID->ProviderData.MessageDatabaseObject.v2Reserved);
									}
								}
	}
		break;
		// Exchange message store folder
	case eidtFolder:
	{
					   Parser.GetWORD(&lpEID->ProviderData.FolderOrMessage.Type);
					   Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.FolderObject.DatabaseGUID),
						   sizeof(lpEID->ProviderData.FolderOrMessage.Data.FolderObject.DatabaseGUID),
						   (LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.FolderObject.DatabaseGUID);
					   Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter),
						   sizeof(lpEID->ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter),
						   (LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter);
					   Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.FolderObject.Pad),
						   sizeof(lpEID->ProviderData.FolderOrMessage.Data.FolderObject.Pad),
						   (LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.FolderObject.Pad);
	}
		break;
		// Exchange message store message
	case eidtMessage:
	{
						Parser.GetWORD(&lpEID->ProviderData.FolderOrMessage.Type);
						Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.FolderDatabaseGUID),
							sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.FolderDatabaseGUID),
							(LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.MessageObject.FolderDatabaseGUID);
						Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter),
							sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter),
							(LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter);
						Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.Pad1),
							sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.Pad1),
							(LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.MessageObject.Pad1);
						Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.MessageDatabaseGUID),
							sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.MessageDatabaseGUID),
							(LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.MessageObject.MessageDatabaseGUID);
						Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter),
							sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter),
							(LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter);
						Parser.GetBYTESNoAlloc(sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.Pad2),
							sizeof(lpEID->ProviderData.FolderOrMessage.Data.MessageObject.Pad2),
							(LPBYTE)&lpEID->ProviderData.FolderOrMessage.Data.MessageObject.Pad2);
	}
		break;
	}

	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();
} // BinToTypedEntryIdStruct

_Check_return_ EntryIdStruct* BinToEntryIdStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	return BinToEntryIdStructWithSize(cbBin, lpBin, NULL);
} // BinToEntryIdStruct

// Allocates return value with new. Clean up with DeleteEntryIdStruct.
// lpcbBytesRead returns the number of bytes consumed
_Check_return_ EntryIdStruct* BinToEntryIdStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead)
{
	if (!lpBin) return NULL;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;

	EntryIdStruct eidEntryId = { 0 };
	CBinaryParser Parser(cbBin, lpBin);
	Parser.GetBYTE(&eidEntryId.abFlags[0]);
	Parser.GetBYTE(&eidEntryId.abFlags[1]);
	Parser.GetBYTE(&eidEntryId.abFlags[2]);
	Parser.GetBYTE(&eidEntryId.abFlags[3]);
	Parser.GetBYTESNoAlloc(sizeof(eidEntryId.ProviderUID), sizeof(eidEntryId.ProviderUID), (LPBYTE)&eidEntryId.ProviderUID);

	// One Off Recipient
	if (!memcmp(eidEntryId.ProviderUID, &muidOOP, sizeof(GUID)))
	{
		eidEntryId.ObjectType = eidtOneOff;
	}
	// Address Book Recipient
	else if (!memcmp(eidEntryId.ProviderUID, &muidEMSAB, sizeof(GUID)))
	{
		eidEntryId.ObjectType = eidtAddressBook;
	}
	// Contact Address Book / Personal Distribution List (PDL)
	else if (!memcmp(eidEntryId.ProviderUID, &muidContabDLL, sizeof(GUID)))
	{
		eidEntryId.ObjectType = eidtContact;
	}
	// message store objects
	else if (!memcmp(eidEntryId.ProviderUID, &muidStoreWrap, sizeof(GUID)))
	{
		eidEntryId.ObjectType = eidtMessageDatabase;
	}
	else if (!memcmp(eidEntryId.ProviderUID, &WAB_GUID, sizeof(GUID)))
	{
		eidEntryId.ObjectType = eidtWAB;
	}
	// We can recognize Exchange message store folder and message entry IDs by their size
	else
	{
		size_t ulRemainingBytes = Parser.RemainingBytes();

		if (sizeof(WORD)+sizeof(GUID)+6 * sizeof(BYTE)+2 * sizeof(BYTE) == ulRemainingBytes)
		{
			eidEntryId.ObjectType = eidtFolder;
		}
		else if (sizeof(WORD)+2 * sizeof(GUID)+12 * sizeof(BYTE)+4 * sizeof(BYTE) == ulRemainingBytes)
		{
			eidEntryId.ObjectType = eidtMessage;
		}
	}

	if (eidtUnknown != eidEntryId.ObjectType)
	{
		size_t cbBinRead = 0;
		BinToTypedEntryIdStruct(
			eidEntryId.ObjectType,
			(ULONG)Parser.RemainingBytes(),
			lpBin + Parser.GetCurrentOffset(),
			&eidEntryId,
			&cbBinRead);
		Parser.Advance(cbBinRead);
	}

	// Check if we have an unidentified short term entry ID:
	if (eidtUnknown == eidEntryId.ObjectType && eidEntryId.abFlags[0] & MAPI_SHORTTERM)
		eidEntryId.ObjectType = eidtShortTerm;

	// Only fill out junk data if we've not been asked to report back how many bytes we read
	// If we've been asked to report back, then someone else will handle the remaining data
	if (!lpcbBytesRead)
	{
		eidEntryId.JunkDataSize = Parser.GetRemainingData(&eidEntryId.JunkData);
	}
	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();

	EntryIdStruct* peidEntryId = new EntryIdStruct;
	if (peidEntryId)
	{
		*peidEntryId = eidEntryId;
	}

	return peidEntryId;
} // BinToEntryIdStructWithSize

void DeleteEntryIdStruct(_In_ EntryIdStruct* peidEntryId)
{
	if (!peidEntryId) return;

	switch (peidEntryId->ObjectType)
	{
	case eidtOneOff:
		if (MAPI_UNICODE & peidEntryId->ProviderData.OneOffRecipientObject.Bitmask)
		{
			delete[] peidEntryId->ProviderData.OneOffRecipientObject.Strings.Unicode.DisplayName;
			delete[] peidEntryId->ProviderData.OneOffRecipientObject.Strings.Unicode.AddressType;
			delete[] peidEntryId->ProviderData.OneOffRecipientObject.Strings.Unicode.EmailAddress;
		}
		else
		{
			delete[] peidEntryId->ProviderData.OneOffRecipientObject.Strings.ANSI.DisplayName;
			delete[] peidEntryId->ProviderData.OneOffRecipientObject.Strings.ANSI.AddressType;
			delete[] peidEntryId->ProviderData.OneOffRecipientObject.Strings.ANSI.EmailAddress;
		}
		break;
	case eidtAddressBook:
		delete[] peidEntryId->ProviderData.AddressBookObject.X500DN;
		break;
	case eidtContact:
		DeleteEntryIdStruct(peidEntryId->ProviderData.ContactAddressBookObject.lpEntryID);
		break;
	case eidtMessageDatabase:
		delete[] peidEntryId->ProviderData.MessageDatabaseObject.DLLFileName;
		delete[] peidEntryId->ProviderData.MessageDatabaseObject.ServerShortname;
		delete[] peidEntryId->ProviderData.MessageDatabaseObject.MailboxDN;
		delete[] peidEntryId->ProviderData.MessageDatabaseObject.v2DN;
		delete[] peidEntryId->ProviderData.MessageDatabaseObject.v2FQDN;
		break;
	case eidtWAB:
		DeleteEntryIdStruct(peidEntryId->ProviderData.WAB.lpEntryID);
	}

	delete[] peidEntryId->JunkData;
	delete peidEntryId;
} // DeleteEntryIdStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR EntryIdStructToString(_In_ EntryIdStruct* peidEntryId)
{
	if (!peidEntryId) return NULL;

	CString szEntryId;
	CString szTmp;

	switch (peidEntryId->ObjectType)
	{
	case eidtUnknown:
		szEntryId.FormatMessage(IDS_ENTRYIDUNKNOWN);
		break;
	case eidtShortTerm:
		szEntryId.FormatMessage(IDS_ENTRYIDSHORTTERM);
		break;
	case eidtFolder:
		szEntryId.FormatMessage(IDS_ENTRYIDEXCHANGEFOLDER);
		break;
	case eidtMessage:
		szEntryId.FormatMessage(IDS_ENTRYIDEXCHANGEMESSAGE);
		break;
	case eidtMessageDatabase:
		szEntryId.FormatMessage(IDS_ENTRYIDMAPIMESSAGESTORE);
		break;
	case eidtOneOff:
		szEntryId.FormatMessage(IDS_ENTRYIDMAPIONEOFF);
		break;
	case eidtAddressBook:
		szEntryId.FormatMessage(IDS_ENTRYIDEXCHANGEADDRESS);
		break;
	case eidtContact:
		szEntryId.FormatMessage(IDS_ENTRYIDCONTACTADDRESS);
		break;
	case eidtWAB:
		szEntryId.FormatMessage(IDS_ENTRYIDWRAPPEDENTRYID);
		break;
	}

	if (0 == (peidEntryId->abFlags[0] | peidEntryId->abFlags[1] | peidEntryId->abFlags[2] | peidEntryId->abFlags[3]))
	{
		szTmp.FormatMessage(IDS_ENTRYIDHEADER1);
	}
	else if (0 == (peidEntryId->abFlags[1] | peidEntryId->abFlags[2] | peidEntryId->abFlags[3]))
	{
		LPTSTR szFlags0 = NULL;
		InterpretFlags(flagEntryId0, peidEntryId->abFlags[0], &szFlags0);
		szTmp.FormatMessage(IDS_ENTRYIDHEADER2,
			peidEntryId->abFlags[0], szFlags0);
		delete[] szFlags0;
		szFlags0 = NULL;
	}
	else
	{
		LPTSTR szFlags0 = NULL;
		LPTSTR szFlags1 = NULL;
		InterpretFlags(flagEntryId0, peidEntryId->abFlags[0], &szFlags0);
		InterpretFlags(flagEntryId1, peidEntryId->abFlags[1], &szFlags1);
		szTmp.FormatMessage(IDS_ENTRYIDHEADER3,
			peidEntryId->abFlags[0], szFlags0,
			peidEntryId->abFlags[1], szFlags1,
			peidEntryId->abFlags[2],
			peidEntryId->abFlags[3]);
		delete[] szFlags1;
		delete[] szFlags0;
		szFlags1 = NULL;
		szFlags0 = NULL;
	}
	szEntryId += szTmp;

	LPTSTR szGUID = NULL;
	szGUID = GUIDToStringAndName((LPGUID)&peidEntryId->ProviderUID);
	szTmp.FormatMessage(IDS_ENTRYIDPROVIDERGUID, szGUID);
	szEntryId += szTmp;

	delete[] szGUID;
	szGUID = NULL;

	if (eidtOneOff == peidEntryId->ObjectType)
	{
		LPTSTR szFlags = NULL;
		InterpretFlags(flagOneOffEntryId, peidEntryId->ProviderData.OneOffRecipientObject.Bitmask, &szFlags);

		if (MAPI_UNICODE & peidEntryId->ProviderData.OneOffRecipientObject.Bitmask)
		{
			szTmp.FormatMessage(IDS_ONEOFFENTRYIDFOOTERUNICODE,
				peidEntryId->ProviderData.OneOffRecipientObject.Bitmask, szFlags,
				peidEntryId->ProviderData.OneOffRecipientObject.Strings.Unicode.DisplayName,
				peidEntryId->ProviderData.OneOffRecipientObject.Strings.Unicode.AddressType,
				peidEntryId->ProviderData.OneOffRecipientObject.Strings.Unicode.EmailAddress);
		}
		else
		{
			szTmp.FormatMessage(IDS_ONEOFFENTRYIDFOOTERANSI,
				peidEntryId->ProviderData.OneOffRecipientObject.Bitmask, szFlags,
				peidEntryId->ProviderData.OneOffRecipientObject.Strings.ANSI.DisplayName,
				peidEntryId->ProviderData.OneOffRecipientObject.Strings.ANSI.AddressType,
				peidEntryId->ProviderData.OneOffRecipientObject.Strings.ANSI.EmailAddress);
		}
		szEntryId += szTmp;
		delete[] szFlags;
		szFlags = NULL;
	}
	else if (eidtAddressBook == peidEntryId->ObjectType)
	{
		LPTSTR szVersion = NULL;
		InterpretFlags(flagExchangeABVersion, peidEntryId->ProviderData.AddressBookObject.Version, &szVersion);
		LPWSTR szType = NULL;
		InterpretNumberAsStringProp(peidEntryId->ProviderData.AddressBookObject.Type, PR_DISPLAY_TYPE, &szType);

		szTmp.FormatMessage(IDS_ENTRYIDEXCHANGEADDRESSDATA,
			peidEntryId->ProviderData.AddressBookObject.Version, szVersion,
			peidEntryId->ProviderData.AddressBookObject.Type, szType,
			peidEntryId->ProviderData.AddressBookObject.X500DN);
		szEntryId += szTmp;

		delete[] szType;
		szType = NULL;
		delete[] szVersion;
		szVersion = NULL;
	}
	// Contact Address Book / Personal Distribution List (PDL)
	else if (eidtContact == peidEntryId->ObjectType)
	{
		LPTSTR szVersion = NULL;
		InterpretFlags(flagContabVersion, peidEntryId->ProviderData.ContactAddressBookObject.Version, &szVersion);
		LPTSTR szType = NULL;
		InterpretFlags(flagContabType, peidEntryId->ProviderData.ContactAddressBookObject.Type, &szType);
		LPTSTR szIndex = NULL;
		InterpretFlags(flagContabIndex, peidEntryId->ProviderData.ContactAddressBookObject.Index, &szIndex);

		szTmp.FormatMessage(IDS_ENTRYIDCONTACTADDRESSDATAHEAD,
			peidEntryId->ProviderData.ContactAddressBookObject.Version, szVersion,
			peidEntryId->ProviderData.ContactAddressBookObject.Type, szType);
		szEntryId += szTmp;

		switch (peidEntryId->ProviderData.ContactAddressBookObject.Type)
		{
		case CONTAB_USER:
		case CONTAB_DISTLIST:
			szTmp.FormatMessage(IDS_ENTRYIDCONTACTADDRESSDATAUSER,
				peidEntryId->ProviderData.ContactAddressBookObject.Index, szIndex,
				peidEntryId->ProviderData.ContactAddressBookObject.EntryIDCount);
			szEntryId += szTmp;
			break;
		case CONTAB_ROOT:
		case CONTAB_CONTAINER:
		case CONTAB_SUBROOT:
			szGUID = GUIDToStringAndName((LPGUID)&peidEntryId->ProviderData.ContactAddressBookObject.muidID);
			szTmp.FormatMessage(IDS_ENTRYIDCONTACTADDRESSDATACONTAINER, szGUID);
			szEntryId += szTmp;

			delete[] szGUID;
			szGUID = NULL;
			break;
		default:
			break;
		}

		delete[] szIndex;
		szIndex = NULL;
		delete[] szType;
		szType = NULL;
		delete[] szVersion;
		szVersion = NULL;

		LPWSTR szEID = EntryIdStructToString(peidEntryId->ProviderData.ContactAddressBookObject.lpEntryID);
		szEntryId += szEID;
		delete[] szEID;
	}
	else if (eidtWAB == peidEntryId->ObjectType)
	{
		LPTSTR szType = NULL;
		InterpretFlags(flagWABEntryIDType, peidEntryId->ProviderData.WAB.Type, &szType);

		szTmp.FormatMessage(IDS_ENTRYIDWRAPPEDENTRYIDDATA, peidEntryId->ProviderData.WAB.Type, szType);
		szEntryId += szTmp;

		LPWSTR szEID = EntryIdStructToString(peidEntryId->ProviderData.WAB.lpEntryID);
		szEntryId += szEID;

		delete[] szType;
		szType = NULL;

		delete[] szEID;
		szEID = NULL;
	}
	else if (eidtMessageDatabase == peidEntryId->ObjectType)
	{
		LPTSTR szVersion = NULL;
		InterpretFlags(flagMDBVersion, peidEntryId->ProviderData.MessageDatabaseObject.Version, &szVersion);

		LPTSTR szFlag = NULL;
		InterpretFlags(flagMDBFlag, peidEntryId->ProviderData.MessageDatabaseObject.Flag, &szFlag);

		szTmp.FormatMessage(IDS_ENTRYIDMAPIMESSAGESTOREDATA,
			peidEntryId->ProviderData.MessageDatabaseObject.Version, szVersion,
			peidEntryId->ProviderData.MessageDatabaseObject.Flag, szFlag,
			peidEntryId->ProviderData.MessageDatabaseObject.DLLFileName);
		szEntryId += szTmp;
		if (peidEntryId->ProviderData.MessageDatabaseObject.bIsExchange)
		{
			LPWSTR szWrappedType = NULL;
			InterpretNumberAsStringProp(peidEntryId->ProviderData.MessageDatabaseObject.WrappedType, PR_PROFILE_OPEN_FLAGS, &szWrappedType);

			szGUID = GUIDToStringAndName((LPGUID)&peidEntryId->ProviderData.MessageDatabaseObject.WrappedProviderUID);
			szTmp.FormatMessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATA,
				peidEntryId->ProviderData.MessageDatabaseObject.WrappedFlags,
				szGUID,
				peidEntryId->ProviderData.MessageDatabaseObject.WrappedType, szWrappedType,
				peidEntryId->ProviderData.MessageDatabaseObject.ServerShortname,
				peidEntryId->ProviderData.MessageDatabaseObject.MailboxDN ? peidEntryId->ProviderData.MessageDatabaseObject.MailboxDN : ""); // STRING_OK
			szEntryId += szTmp;

			delete[] szGUID;
			szGUID = NULL;

			delete[] szWrappedType;
			szWrappedType = NULL;
		}

		if (peidEntryId->ProviderData.MessageDatabaseObject.bV2)
		{
			LPTSTR szV2Magic = NULL;
			LPTSTR szV2Version = NULL;
			InterpretFlags(flagV2Magic, peidEntryId->ProviderData.MessageDatabaseObject.v2.ulMagic, &szV2Magic);
			InterpretFlags(flagV2Version, peidEntryId->ProviderData.MessageDatabaseObject.v2.ulVersion, &szV2Version);

			szTmp.FormatMessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATAV2,
				peidEntryId->ProviderData.MessageDatabaseObject.v2.ulMagic, szV2Magic,
				peidEntryId->ProviderData.MessageDatabaseObject.v2.ulSize,
				peidEntryId->ProviderData.MessageDatabaseObject.v2.ulVersion, szV2Version,
				peidEntryId->ProviderData.MessageDatabaseObject.v2.ulOffsetDN,
				peidEntryId->ProviderData.MessageDatabaseObject.v2.ulOffsetFQDN,
				peidEntryId->ProviderData.MessageDatabaseObject.v2DN ? peidEntryId->ProviderData.MessageDatabaseObject.v2DN : "", // STRING_OK
				peidEntryId->ProviderData.MessageDatabaseObject.v2FQDN ? peidEntryId->ProviderData.MessageDatabaseObject.v2FQDN : L""); // STRING_OK
			szEntryId += szTmp;

			delete[] szV2Magic;
			szV2Magic = NULL;

			delete[] szV2Version;
			szV2Version = NULL;

			SBinary sBin = { 0 };
			sBin.cb = sizeof(peidEntryId->ProviderData.MessageDatabaseObject.v2Reserved);
			sBin.lpb = peidEntryId->ProviderData.MessageDatabaseObject.v2Reserved;
			szEntryId += wstringToCString(BinToHexString(&sBin, true));
		}

		delete[] szFlag;
		szFlag = NULL;

		delete[] szVersion;
		szVersion = NULL;
	}
	else if (eidtFolder == peidEntryId->ObjectType)
	{
		LPTSTR szType = NULL;
		InterpretFlags(flagMessageDatabaseObjectType, peidEntryId->ProviderData.FolderOrMessage.Type, &szType);

		LPTSTR szDatabaseGUID = GUIDToStringAndName((LPGUID)&peidEntryId->ProviderData.FolderOrMessage.Data.FolderObject.DatabaseGUID);

		SBinary sBinGlobalCounter = { 0 };
		sBinGlobalCounter.cb = sizeof(peidEntryId->ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter);
		sBinGlobalCounter.lpb = peidEntryId->ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter;

		SBinary sBinPad = { 0 };
		sBinPad.cb = sizeof(peidEntryId->ProviderData.FolderOrMessage.Data.FolderObject.Pad);
		sBinPad.lpb = peidEntryId->ProviderData.FolderOrMessage.Data.FolderObject.Pad;

		szTmp.FormatMessage(IDS_ENTRYIDEXCHANGEFOLDERDATA,
			peidEntryId->ProviderData.FolderOrMessage.Type, szType,
			szDatabaseGUID);
		szEntryId += szTmp;
		szEntryId += wstringToCString(BinToHexString(&sBinGlobalCounter, true));

		szTmp.FormatMessage(IDS_ENTRYIDEXCHANGEDATAPAD);
		szEntryId += szTmp;
		szEntryId += wstringToCString(BinToHexString(&sBinPad, true));

		delete[] szDatabaseGUID;
		szDatabaseGUID = NULL;
		delete[] szType;
		szType = NULL;
	}
	else if (eidtMessage == peidEntryId->ObjectType)
	{
		LPTSTR szType = NULL;
		InterpretFlags(flagMessageDatabaseObjectType, peidEntryId->ProviderData.FolderOrMessage.Type, &szType);

		LPTSTR szFolderDatabaseGUID = GUIDToStringAndName((LPGUID)&peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.FolderDatabaseGUID);

		SBinary sBinFolderGlobalCounter = { 0 };
		sBinFolderGlobalCounter.cb = sizeof(peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter);
		sBinFolderGlobalCounter.lpb = peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter;

		SBinary sBinPad1 = { 0 };
		sBinPad1.cb = sizeof(peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.Pad1);
		sBinPad1.lpb = peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.Pad1;

		LPTSTR szMessageDatabaseGUID = GUIDToStringAndName((LPGUID)&peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.MessageDatabaseGUID);

		SBinary sBinMessageGlobalCounter = { 0 };
		sBinMessageGlobalCounter.cb = sizeof(peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter);
		sBinMessageGlobalCounter.lpb = peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter;

		SBinary sBinPad2 = { 0 };
		sBinPad2.cb = sizeof(peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.Pad2);
		sBinPad2.lpb = peidEntryId->ProviderData.FolderOrMessage.Data.MessageObject.Pad2;

		szTmp.FormatMessage(IDS_ENTRYIDEXCHANGEMESSAGEDATA,
			peidEntryId->ProviderData.FolderOrMessage.Type, szType,
			szFolderDatabaseGUID);
		szEntryId += szTmp;
		szEntryId += wstringToCString(BinToHexString(&sBinFolderGlobalCounter, true));

		szTmp.FormatMessage(IDS_ENTRYIDEXCHANGEDATAPADNUM, 1);
		szEntryId += szTmp;
		szEntryId += wstringToCString(BinToHexString(&sBinPad1, true));

		szTmp.FormatMessage(IDS_ENTRYIDEXCHANGEMESSAGEDATAGUID,
			szMessageDatabaseGUID);
		szEntryId += szTmp;
		szEntryId += wstringToCString(BinToHexString(&sBinMessageGlobalCounter, true));

		szTmp.FormatMessage(IDS_ENTRYIDEXCHANGEDATAPADNUM, 2);
		szEntryId += szTmp;
		szEntryId += wstringToCString(BinToHexString(&sBinPad2, true));

		delete[] szMessageDatabaseGUID;
		szMessageDatabaseGUID = NULL;

		delete[] szFolderDatabaseGUID;
		szFolderDatabaseGUID = NULL;

		delete[] szType;
		szType = NULL;
	}

	szEntryId += JunkDataToString(peidEntryId->JunkDataSize, peidEntryId->JunkData);

	return CStringToLPWSTR(szEntryId);
} // EntryIdStructToString

//////////////////////////////////////////////////////////////////////////
// End BinToEntryIdStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// PropertyStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new. Clean up with DeletePropertyStruct.
_Check_return_ PropertyStruct* BinToPropertyStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;
	CBinaryParser Parser(cbBin, lpBin);

	size_t cbBytesRead = 0;

	// Have to count how many properties are here.
	// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
	size_t stBookmark = Parser.GetCurrentOffset();

	DWORD dwPropCount = 0;

	for (;;)
	{
		LPSPropValue lpProp = NULL;
		lpProp = BinToSPropValue(
			(ULONG)Parser.RemainingBytes(),
			lpBin + Parser.GetCurrentOffset(),
			1,
			&cbBytesRead,
			false);
		DeleteSPropVal(1, lpProp);
		if (!cbBytesRead) break; // we hit the last, zero byte block, or the end of the buffer
		dwPropCount++;
		Parser.Advance(cbBytesRead);
	}
	Parser.SetCurrentOffset(stBookmark); // We're done with our first pass, restore the bookmark

	PropertyStruct* ppProperty = new PropertyStruct;
	if (ppProperty)
	{
		memset(ppProperty, 0, sizeof(PropertyStruct));
		ppProperty->PropCount = dwPropCount;
		ppProperty->Prop = BinToSPropValue(
			cbBin,
			lpBin,
			dwPropCount,
			&cbBytesRead,
			false);
		Parser.Advance(cbBytesRead);

		ppProperty->JunkDataSize = Parser.GetRemainingData(&ppProperty->JunkData);
	}

	return ppProperty;
} // BinToPropertyStruct

// Caller allocates with new. Clean up with DeleteSPropVal.
_Check_return_ LPSPropValue BinToSPropValue(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, DWORD dwPropCount, _Out_ size_t* lpcbBytesRead, bool bStringPropsExcludeLength)
{
	if (!lpBin) return NULL;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	if (!dwPropCount || dwPropCount > _MaxEntriesSmall) return NULL;
	LPSPropValue pspvProperty = new SPropValue[dwPropCount];
	if (!pspvProperty) return NULL;
	memset(pspvProperty, 0, sizeof(SPropValue)*dwPropCount);
	CBinaryParser Parser(cbBin, lpBin);

	DWORD i = 0;

	for (i = 0; i < dwPropCount; i++)
	{
		WORD PropType = 0;
		WORD PropID = 0;

		Parser.GetWORD(&PropType);
		Parser.GetWORD(&PropID);

		pspvProperty[i].ulPropTag = PROP_TAG(PropType, PropID);
		pspvProperty[i].dwAlignPad = 0;

		DWORD dwTemp = 0;
		WORD wTemp = 0;
		switch (PropType)
		{
		case PT_LONG:
			Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.l = dwTemp;
			break;
		case PT_ERROR:
			Parser.GetDWORD(&dwTemp);
			pspvProperty[i].Value.err = dwTemp;
			break;
		case PT_BOOLEAN:
			Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.b = wTemp;
			break;
		case PT_UNICODE:
			if (bStringPropsExcludeLength)
			{
				Parser.GetStringW(&pspvProperty[i].Value.lpszW);
			}
			else
			{
				// This is apparently a cb...
				Parser.GetWORD(&wTemp);
				Parser.GetStringW(wTemp / sizeof(WCHAR), &pspvProperty[i].Value.lpszW);
			}
			break;
		case PT_STRING8:
			if (bStringPropsExcludeLength)
			{
				Parser.GetStringA(&pspvProperty[i].Value.lpszA);
			}
			else
			{
				// This is apparently a cb...
				Parser.GetWORD(&wTemp);
				Parser.GetStringA(wTemp, &pspvProperty[i].Value.lpszA);
			}
			break;
		case PT_SYSTIME:
			Parser.GetDWORD(&pspvProperty[i].Value.ft.dwHighDateTime);
			Parser.GetDWORD(&pspvProperty[i].Value.ft.dwLowDateTime);
			break;
		case PT_BINARY:
			Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.bin.cb = wTemp;
			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			Parser.GetBYTES(pspvProperty[i].Value.bin.cb, pspvProperty[i].Value.bin.cb, &pspvProperty[i].Value.bin.lpb);
			break;
		case PT_MV_STRING8:
			Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.MVszA.cValues = wTemp;
			pspvProperty[i].Value.MVszA.lppszA = new CHAR*[wTemp];
			if (pspvProperty[i].Value.MVszA.lppszA)
			{
				memset(pspvProperty[i].Value.MVszA.lppszA, 0, sizeof(CHAR*)* wTemp);
				DWORD j = 0;
				for (j = 0; j < pspvProperty[i].Value.MVszA.cValues; j++)
				{
					Parser.GetStringA(&pspvProperty[i].Value.MVszA.lppszA[j]);
				}
			}
			break;
		case PT_MV_UNICODE:
			Parser.GetWORD(&wTemp);
			pspvProperty[i].Value.MVszW.cValues = wTemp;
			pspvProperty[i].Value.MVszW.lppszW = new WCHAR*[wTemp];
			if (pspvProperty[i].Value.MVszW.lppszW)
			{
				memset(pspvProperty[i].Value.MVszW.lppszW, 0, sizeof(WCHAR*)* wTemp);
				DWORD j = 0;
				for (j = 0; j < pspvProperty[i].Value.MVszW.cValues; j++)
				{
					Parser.GetStringW(&pspvProperty[i].Value.MVszW.lppszW[j]);
				}
			}
			break;
		default:
			break;
		}
	}

	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();
	return pspvProperty;
} // BinToSPropValue

void DeleteSPropVal(ULONG cVal, _In_count_(cVal) LPSPropValue lpsPropVal)
{
	if (!lpsPropVal) return;
	DWORD i = 0;
	DWORD j = 0;
	for (i = 0; i < cVal; i++)
	{
		switch (PROP_TYPE(lpsPropVal[i].ulPropTag))
		{
		case PT_UNICODE:
			delete[] lpsPropVal[i].Value.lpszW;
			break;
		case PT_STRING8:
			delete[] lpsPropVal[i].Value.lpszA;
			break;
		case PT_BINARY:
			delete[] lpsPropVal[i].Value.bin.lpb;
			break;
		case PT_MV_STRING8:
			if (lpsPropVal[i].Value.MVszA.lppszA)
			{
				for (j = 0; j < lpsPropVal[i].Value.MVszA.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVszA.lppszA[j];
				}
			}
			break;
		case PT_MV_UNICODE:
			if (lpsPropVal[i].Value.MVszW.lppszW)
			{
				for (j = 0; j < lpsPropVal[i].Value.MVszW.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVszW.lppszW[j];
				}
			}
			break;
		case PT_MV_BINARY:
			if (lpsPropVal[i].Value.MVbin.lpbin)
			{
				for (j = 0; j < lpsPropVal[i].Value.MVbin.cValues; j++)
				{
					delete[] lpsPropVal[i].Value.MVbin.lpbin[j].lpb;
				}
			}
			break;
		default:
			break;
		}
	}
	delete[] lpsPropVal;
} // DeleteSPropVal

void DeletePropertyStruct(_In_ PropertyStruct* ppProperty)
{
	if (!ppProperty) return;

	DeleteSPropVal(ppProperty->PropCount, ppProperty->Prop);

	delete[] ppProperty->JunkData;
	delete ppProperty;
} // DeletePropertyStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR PropertyStructToString(_In_ PropertyStruct* ppProperty)
{
	if (!ppProperty) return NULL;

	CString szProperty;
	CString szTmp;
	DWORD i = 0;

	if (ppProperty->Prop)
	{
		for (i = 0; i < ppProperty->PropCount; i++)
		{
			wstring PropString;
			wstring AltPropString;
			LPTSTR szExactMatches = NULL;
			LPTSTR szPartialMatches = NULL;
			LPWSTR szSmartView = NULL;

			InterpretPropSmartView(
				&ppProperty->Prop[i],
				NULL,
				NULL,
				NULL,
				false,
				false,
				&szSmartView);

			szTmp.FormatMessage(IDS_PROPERTYDATAHEADER,
				i,
				ppProperty->Prop[i].ulPropTag);
			szProperty += szTmp;

			PropTagToPropName(ppProperty->Prop[i].ulPropTag, false, &szExactMatches, &szPartialMatches);

			if (!IsNullOrEmpty(szExactMatches))
			{
				szTmp.FormatMessage(IDS_PROPERTYDATAEXACTMATCHES,
					szExactMatches);
				szProperty += szTmp;
			}

			if (!IsNullOrEmpty(szPartialMatches))
			{
				szTmp.FormatMessage(IDS_PROPERTYDATAPARTIALMATCHES,
					szPartialMatches);
				szProperty += szTmp;
			}

			InterpretProp(&ppProperty->Prop[i], &PropString, &AltPropString);
			szTmp.FormatMessage(IDS_PROPERTYDATA,
				PropString.c_str(),
				AltPropString.c_str());
			szProperty += szTmp;

			if (szSmartView)
			{
				szTmp.FormatMessage(IDS_PROPERTYDATASMARTVIEW,
					szSmartView);
				szProperty += szTmp;
			}

			delete[] szExactMatches;
			delete[] szPartialMatches;
			delete[] szSmartView;
		}
	}

	szProperty += JunkDataToString(ppProperty->JunkDataSize, ppProperty->JunkData);

	return CStringToLPWSTR(szProperty);
}

//////////////////////////////////////////////////////////////////////////
// End PropertyStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// RestrictionStruct
//////////////////////////////////////////////////////////////////////////

_Check_return_ RestrictionStruct* BinToRestrictionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	return BinToRestrictionStructWithSize(cbBin, lpBin, NULL);
} // BinToRestrictionStruct

// Allocates return value with new. Clean up with DeleteRestrictionStruct.
_Check_return_ RestrictionStruct* BinToRestrictionStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead)
{
	if (!lpBin) return NULL;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	CBinaryParser Parser(cbBin, lpBin);

	size_t cbBytesRead = 0;

	RestrictionStruct* prRestriction = new RestrictionStruct;
	if (prRestriction)
	{
		memset(prRestriction, 0, sizeof(RestrictionStruct));
		prRestriction->lpRes = new SRestriction;

		if (prRestriction->lpRes)
		{
			memset(prRestriction->lpRes, 0, sizeof(SRestriction));

			(void)BinToRestriction(
				0,
				cbBin,
				lpBin,
				&cbBytesRead,
				prRestriction->lpRes,
				false,
				true);
			Parser.Advance(cbBytesRead);
		}

		// Only fill out junk data if we've not been asked to report back how many bytes we read
		// If we've been asked to report back, then someone else will handle the remaining data
		if (!lpcbBytesRead)
		{
			prRestriction->JunkDataSize = Parser.GetRemainingData(&prRestriction->JunkData);
		}
	}
	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();

	return prRestriction;
} // BinToRestrictionStructWithSize

// Helper function for both RestrictionStruct and RuleConditionStruct
// Caller allocates with new. Clean up with DeleteRestriction and delete[].
// If bRuleCondition is true, parse restrictions defined in [MS-OXCDATA] 2.13
// If bRuleCondition is true, bExtendedCount controls whether the count fields in AND/OR restrictions is 16 or 32 bits
// If bRuleCondition is false, parse restrictions defined in [MS-OXOCFG] 2.2.4.1.2
// If bRuleCondition is false, ignore bExtendedCount (assumes true)
// Returns true if it successfully read a restriction
bool BinToRestriction(ULONG ulDepth, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _In_ LPSRestriction psrRestriction, bool bRuleCondition, bool bExtendedCount)
{
	if (ulDepth > _MaxDepth) return false;
	if (!lpBin) return false;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	if (!psrRestriction) return false;
	CBinaryParser Parser(cbBin, lpBin);

	BYTE bTemp = 0;
	WORD wTemp = 0;
	DWORD dwTemp = 0;
	DWORD i = 0;
	size_t cbOffset = 0;
	size_t cbBytesRead = 0;
	bool bRet = true;

	if (bRuleCondition)
	{
		Parser.GetBYTE(&bTemp);
		psrRestriction->rt = bTemp;
	}
	else
	{
		Parser.GetDWORD(&dwTemp);
		psrRestriction->rt = dwTemp;
	}

	switch (psrRestriction->rt)
	{
	case RES_AND:
	case RES_OR:
		if (!bRuleCondition || bExtendedCount)
		{
			Parser.GetDWORD(&dwTemp);
		}
		else
		{
			Parser.GetWORD(&wTemp);
			dwTemp = wTemp;
		}
		psrRestriction->res.resAnd.cRes = dwTemp;
		if (psrRestriction->res.resAnd.cRes &&
			psrRestriction->res.resAnd.cRes < _MaxEntriesExtraLarge)
		{
			psrRestriction->res.resAnd.lpRes = new SRestriction[psrRestriction->res.resAnd.cRes];
			if (psrRestriction->res.resAnd.lpRes)
			{
				memset(psrRestriction->res.resAnd.lpRes, 0, sizeof(SRestriction)* psrRestriction->res.resAnd.cRes);
				for (i = 0; i < psrRestriction->res.resAnd.cRes; i++)
				{
					cbOffset = Parser.GetCurrentOffset();
					bRet = BinToRestriction(
						ulDepth + 1,
						(ULONG)Parser.RemainingBytes(),
						lpBin + cbOffset,
						&cbBytesRead,
						&psrRestriction->res.resAnd.lpRes[i],
						bRuleCondition,
						bExtendedCount);
					Parser.Advance(cbBytesRead);
					if (!bRet) break;
				}
			}
		}
		break;
	case RES_NOT:
		psrRestriction->res.resNot.lpRes = new SRestriction;
		if (psrRestriction->res.resNot.lpRes)
		{
			memset(psrRestriction->res.resNot.lpRes, 0, sizeof(SRestriction));
			cbOffset = Parser.GetCurrentOffset();
			bRet = BinToRestriction(
				ulDepth + 1,
				(ULONG)Parser.RemainingBytes(),
				lpBin + cbOffset,
				&cbBytesRead,
				psrRestriction->res.resNot.lpRes,
				bRuleCondition,
				bExtendedCount);
			Parser.Advance(cbBytesRead);
		}
		break;
	case RES_CONTENT:
		Parser.GetDWORD(&psrRestriction->res.resContent.ulFuzzyLevel);
		Parser.GetDWORD(&psrRestriction->res.resContent.ulPropTag);
		cbOffset = Parser.GetCurrentOffset();
		psrRestriction->res.resContent.lpProp = BinToSPropValue(
			(ULONG)Parser.RemainingBytes(),
			lpBin + cbOffset,
			1,
			&cbBytesRead,
			bRuleCondition);
		Parser.Advance(cbBytesRead);
		break;
	case RES_PROPERTY:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resProperty.relop);
		else
			Parser.GetDWORD(&psrRestriction->res.resProperty.relop);
		Parser.GetDWORD(&psrRestriction->res.resProperty.ulPropTag);
		cbOffset = Parser.GetCurrentOffset();
		psrRestriction->res.resProperty.lpProp = BinToSPropValue(
			(ULONG)Parser.RemainingBytes(),
			lpBin + cbOffset,
			1,
			&cbBytesRead,
			bRuleCondition);
		Parser.Advance(cbBytesRead);
		break;
	case RES_COMPAREPROPS:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resCompareProps.relop);
		else
			Parser.GetDWORD(&psrRestriction->res.resCompareProps.relop);
		Parser.GetDWORD(&psrRestriction->res.resCompareProps.ulPropTag1);
		Parser.GetDWORD(&psrRestriction->res.resCompareProps.ulPropTag2);
		break;
	case RES_BITMASK:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resBitMask.relBMR);
		else
			Parser.GetDWORD(&psrRestriction->res.resBitMask.relBMR);
		Parser.GetDWORD(&psrRestriction->res.resBitMask.ulPropTag);
		Parser.GetDWORD(&psrRestriction->res.resBitMask.ulMask);
		break;
	case RES_SIZE:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resSize.relop);
		else
			Parser.GetDWORD(&psrRestriction->res.resSize.relop);
		Parser.GetDWORD(&psrRestriction->res.resSize.ulPropTag);
		Parser.GetDWORD(&psrRestriction->res.resSize.cb);
		break;
	case RES_EXIST:
		Parser.GetDWORD(&psrRestriction->res.resExist.ulPropTag);
		break;
	case RES_SUBRESTRICTION:
		Parser.GetDWORD(&psrRestriction->res.resSub.ulSubObject);
		psrRestriction->res.resSub.lpRes = new SRestriction;
		if (psrRestriction->res.resSub.lpRes)
		{
			memset(psrRestriction->res.resSub.lpRes, 0, sizeof(SRestriction));
			cbOffset = Parser.GetCurrentOffset();
			bRet = BinToRestriction(
				ulDepth + 1,
				(ULONG)Parser.RemainingBytes(),
				lpBin + cbOffset,
				&cbBytesRead,
				psrRestriction->res.resSub.lpRes,
				bRuleCondition,
				bExtendedCount);
			Parser.Advance(cbBytesRead);
		}
		break;
	case RES_COMMENT:
		if (bRuleCondition)
			Parser.GetBYTE((LPBYTE)&psrRestriction->res.resComment.cValues);
		else
			Parser.GetDWORD(&psrRestriction->res.resComment.cValues);
		cbOffset = Parser.GetCurrentOffset();
		psrRestriction->res.resProperty.lpProp = BinToSPropValue(
			(ULONG)Parser.RemainingBytes(),
			lpBin + cbOffset,
			psrRestriction->res.resComment.cValues,
			&cbBytesRead,
			bRuleCondition);
		Parser.Advance(cbBytesRead);

		// Check if a restriction is present
		Parser.GetBYTE(&bTemp);
		if (bTemp)
		{
			psrRestriction->res.resComment.lpRes = new SRestriction;
			if (psrRestriction->res.resComment.lpRes)
			{
				memset(psrRestriction->res.resComment.lpRes, 0, sizeof(SRestriction));
				cbOffset = Parser.GetCurrentOffset();
				bRet = BinToRestriction(
					ulDepth + 1,
					(ULONG)Parser.RemainingBytes(),
					lpBin + cbOffset,
					&cbBytesRead,
					psrRestriction->res.resComment.lpRes,
					bRuleCondition,
					bExtendedCount);
				Parser.Advance(cbBytesRead);
			}
		}
		break;
	case RES_COUNT:
		// RES_COUNT and RES_NOT look the same, so we use the resNot member here
		Parser.GetDWORD(&psrRestriction->res.resNot.ulReserved);
		psrRestriction->res.resNot.lpRes = new SRestriction;
		if (psrRestriction->res.resNot.lpRes)
		{
			memset(psrRestriction->res.resNot.lpRes, 0, sizeof(SRestriction));
			cbOffset = Parser.GetCurrentOffset();
			bRet = BinToRestriction(
				ulDepth + 1,
				(ULONG)Parser.RemainingBytes(),
				lpBin + cbOffset,
				&cbBytesRead,
				psrRestriction->res.resNot.lpRes,
				bRuleCondition,
				bExtendedCount);
			Parser.Advance(cbBytesRead);
		}
		break;
	}

	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();
	return bRet;
} // BinToRestriction

// Does not delete lpRes, which must be released manually. See RES_AND case below.
void DeleteRestriction(_In_ LPSRestriction lpRes)
{
	if (!lpRes) return;

	switch (lpRes->rt)
	{
	case RES_AND:
	case RES_OR:
		if (lpRes->res.resAnd.lpRes)
		{
			DWORD i = 0;
			for (i = 0; i < lpRes->res.resAnd.cRes; i++)
			{
				DeleteRestriction(&lpRes->res.resAnd.lpRes[i]);
			}
		}
		delete[] lpRes->res.resAnd.lpRes;
		break;
		// Structures for these two types are identical
	case RES_NOT:
	case RES_COUNT:
		DeleteRestriction(lpRes->res.resNot.lpRes);
		delete[] lpRes->res.resNot.lpRes;
		break;
	case RES_CONTENT:
		DeleteSPropVal(1, lpRes->res.resContent.lpProp);
		break;
	case RES_PROPERTY:
		DeleteSPropVal(1, lpRes->res.resProperty.lpProp);
		break;
	case RES_SUBRESTRICTION:
		DeleteRestriction(lpRes->res.resSub.lpRes);
		delete[] lpRes->res.resSub.lpRes;
		break;
	case RES_COMMENT:
		if (lpRes->res.resComment.cValues)
		{
			DeleteSPropVal(lpRes->res.resComment.cValues, lpRes->res.resComment.lpProp);
		}
		DeleteRestriction(lpRes->res.resComment.lpRes);
		delete[] lpRes->res.resComment.lpRes;
		break;
	}
} // DeleteRestriction

// Neuters the Restriction struct - caller must use delete to delete the struct itself
void DeleteRestrictionStruct(_In_ RestrictionStruct* prRestriction)
{
	if (!prRestriction) return;

	DeleteRestriction(prRestriction->lpRes);
	delete[] prRestriction->lpRes;

	delete[] prRestriction->JunkData;
	delete prRestriction;
} // DeleteRestrictionStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR RestrictionStructToString(_In_ RestrictionStruct* prRestriction)
{
	if (!prRestriction) return NULL;

	CString szRestriction;
	CString szTmp;

	szRestriction.FormatMessage(IDS_RESTRICTIONDATA);
	szRestriction += RestrictionToString(prRestriction->lpRes, NULL);

	szRestriction += JunkDataToString(prRestriction->JunkDataSize, prRestriction->JunkData);

	return CStringToLPWSTR(szRestriction);
} // RestrictionStructToString

//////////////////////////////////////////////////////////////////////////
// End RestrictionStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// RuleConditionStruct
//////////////////////////////////////////////////////////////////////////

// Rule Condition - these are used in Rules messages
// bExtended indicates an Extended Rules Message
void RuleConditionToString(SBinary myBin, _Deref_out_opt_z_ LPWSTR* lpszResultString, bool bExtended)
{
	if (!lpszResultString) return;
	*lpszResultString = NULL;
	RuleConditionStruct* prcRuleCondition = BinToRuleConditionStruct(myBin.cb, myBin.lpb, NULL, bExtended);

	if (prcRuleCondition)
	{
		*lpszResultString = RuleConditionStructToString(prcRuleCondition, bExtended);
		DeleteRuleConditionStruct(prcRuleCondition);
	}
} // RuleConditionToString

// Allocates return value with new. Clean up with DeleteRuleConditionStruct.
_Check_return_ RuleConditionStruct* BinToRuleConditionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead, bool bExtended)
{
	if (!lpBin) return NULL;
	if (lpcbBytesRead) *lpcbBytesRead = NULL;
	CBinaryParser Parser(cbBin, lpBin);

	size_t cbBytesRead = 0;

	RuleConditionStruct* prcRuleCondition = new RuleConditionStruct;
	if (prcRuleCondition)
	{
		memset(prcRuleCondition, 0, sizeof(RuleConditionStruct));
		Parser.GetWORD(&prcRuleCondition->NamedPropertyInformation.NoOfNamedProps);
		if (prcRuleCondition->NamedPropertyInformation.NoOfNamedProps && prcRuleCondition->NamedPropertyInformation.NoOfNamedProps < _MaxEntriesLarge)
		{
			prcRuleCondition->NamedPropertyInformation.PropId = new WORD[prcRuleCondition->NamedPropertyInformation.NoOfNamedProps];
			if (prcRuleCondition->NamedPropertyInformation.PropId)
			{
				memset(prcRuleCondition->NamedPropertyInformation.PropId, 0, prcRuleCondition->NamedPropertyInformation.NoOfNamedProps*sizeof(WORD));
				WORD i = 0;
				for (i = 0; i < prcRuleCondition->NamedPropertyInformation.NoOfNamedProps; i++)
				{
					Parser.GetWORD(&prcRuleCondition->NamedPropertyInformation.PropId[i]);
				}
			}
			Parser.GetDWORD(&prcRuleCondition->NamedPropertyInformation.NamedPropertiesSize);
			prcRuleCondition->NamedPropertyInformation.PropertyName = new PropertyNameStruct[prcRuleCondition->NamedPropertyInformation.NoOfNamedProps];
			if (prcRuleCondition->NamedPropertyInformation.PropertyName)
			{
				memset(prcRuleCondition->NamedPropertyInformation.PropertyName, 0, prcRuleCondition->NamedPropertyInformation.NoOfNamedProps*sizeof(PropertyNameStruct));
				WORD i = 0;
				for (i = 0; i < prcRuleCondition->NamedPropertyInformation.NoOfNamedProps; i++)
				{
					Parser.GetBYTE(&prcRuleCondition->NamedPropertyInformation.PropertyName[i].Kind);
					Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)&prcRuleCondition->NamedPropertyInformation.PropertyName[i].Guid);
					if (MNID_ID == prcRuleCondition->NamedPropertyInformation.PropertyName[i].Kind)
					{
						Parser.GetDWORD(&prcRuleCondition->NamedPropertyInformation.PropertyName[i].LID);
					}
					else if (MNID_STRING == prcRuleCondition->NamedPropertyInformation.PropertyName[i].Kind)
					{
						Parser.GetBYTE(&prcRuleCondition->NamedPropertyInformation.PropertyName[i].NameSize);
						Parser.GetStringW(
							prcRuleCondition->NamedPropertyInformation.PropertyName[i].NameSize / sizeof(WCHAR),
							&prcRuleCondition->NamedPropertyInformation.PropertyName[i].Name);
					}
				}
			}
		}

		prcRuleCondition->lpRes = new SRestriction;

		if (prcRuleCondition->lpRes)
		{
			memset(prcRuleCondition->lpRes, 0, sizeof(SRestriction));

			(void)BinToRestriction(
				0,
				(ULONG)Parser.RemainingBytes(),
				lpBin + Parser.GetCurrentOffset(),
				&cbBytesRead,
				prcRuleCondition->lpRes,
				true,
				bExtended);
			Parser.Advance(cbBytesRead);
		}

		// Only fill out junk data if we've not been asked to report back how many bytes we read
		// If we've been asked to report back, then someone else will handle the remaining data
		if (!lpcbBytesRead)
		{
			prcRuleCondition->JunkDataSize = Parser.GetRemainingData(&prcRuleCondition->JunkData);
		}
	}
	if (lpcbBytesRead) *lpcbBytesRead = Parser.GetCurrentOffset();

	return prcRuleCondition;
} // BinToRuleConditionStruct

void DeleteRuleConditionStruct(_In_ RuleConditionStruct* prcRuleCondition)
{
	if (!prcRuleCondition) return;

	delete[] prcRuleCondition->NamedPropertyInformation.PropId;
	WORD i = 0;
	prcRuleCondition->NamedPropertyInformation.NoOfNamedProps;
	if (prcRuleCondition->NamedPropertyInformation.PropertyName)
	{
		for (i = 0; i < prcRuleCondition->NamedPropertyInformation.NoOfNamedProps; i++)
		{
			delete[] prcRuleCondition->NamedPropertyInformation.PropertyName[i].Name;
		}
	}
	delete[] prcRuleCondition->NamedPropertyInformation.PropertyName;
	DeleteRestriction(prcRuleCondition->lpRes);
	delete[] prcRuleCondition->lpRes;

	delete[] prcRuleCondition->JunkData;
	delete[] prcRuleCondition;
} // DeleteRuleConditionStruct

_Check_return_ LPWSTR RuleConditionStructToString(_In_ RuleConditionStruct* prcRuleCondition, bool bExtended)
{
	if (!prcRuleCondition) return NULL;

	CString szRuleCondition;
	CString szRestriction;
	CString szTmp;
	if (bExtended)
	{
		szRuleCondition.FormatMessage(IDS_EXRULECONHEADER,
			prcRuleCondition->NamedPropertyInformation.NoOfNamedProps);
	}
	else
	{
		szRuleCondition.FormatMessage(IDS_RULECONHEADER,
			prcRuleCondition->NamedPropertyInformation.NoOfNamedProps);
	}
	if (prcRuleCondition->NamedPropertyInformation.NoOfNamedProps && prcRuleCondition->NamedPropertyInformation.PropId)
	{
		szTmp.FormatMessage(IDS_RULECONNAMEPROPSIZE,
			prcRuleCondition->NamedPropertyInformation.NamedPropertiesSize);
		szRuleCondition += szTmp;

		WORD i = 0;
		for (i = 0; i < prcRuleCondition->NamedPropertyInformation.NoOfNamedProps; i++)
		{
			szTmp.FormatMessage(IDS_RULECONNAMEPROPID, i, prcRuleCondition->NamedPropertyInformation.PropId[i]);
			szRuleCondition += szTmp;

			szTmp.FormatMessage(IDS_RULECONNAMEPROPKIND,
				prcRuleCondition->NamedPropertyInformation.PropertyName[i].Kind);
			szRuleCondition += szTmp;

			LPTSTR szGUID = GUIDToString(&prcRuleCondition->NamedPropertyInformation.PropertyName[i].Guid);
			szRuleCondition += szGUID;
			delete[] szGUID;
			szGUID = NULL;

			if (MNID_ID == prcRuleCondition->NamedPropertyInformation.PropertyName[i].Kind)
			{
				szTmp.FormatMessage(IDS_RULECONNAMEPROPLID,
					prcRuleCondition->NamedPropertyInformation.PropertyName[i].LID);
				szRuleCondition += szTmp;
			}
			else if (MNID_STRING == prcRuleCondition->NamedPropertyInformation.PropertyName[i].Kind)
			{
				szTmp.FormatMessage(IDS_RULENAMEPROPSIZE,
					prcRuleCondition->NamedPropertyInformation.PropertyName[i].NameSize);
				szRuleCondition += szTmp;
				if (prcRuleCondition->NamedPropertyInformation.PropertyName[i].Name)
				{
					szRuleCondition += prcRuleCondition->NamedPropertyInformation.PropertyName[i].Name;
				}
			}
		}
	}

	szRuleCondition += _T("\r\n"); // STRING_OK
	szRestriction.FormatMessage(IDS_RESTRICTIONDATA);
	szRestriction += RestrictionToString(prcRuleCondition->lpRes, NULL);
	szRuleCondition += szRestriction;

	szRuleCondition += JunkDataToString(prcRuleCondition->JunkDataSize, prcRuleCondition->JunkData);

	return CStringToLPWSTR(szRuleCondition);
} // RuleConditionStructToString

//////////////////////////////////////////////////////////////////////////
// End RuleConditionStruct
//////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////
// EntryListStruct
//////////////////////////////////////////////////////////////////////////

// Allocates return value with new. Clean up with DeleteEntryListStruct.
_Check_return_ EntryListStruct* BinToEntryListStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	if (!lpBin) return NULL;

	EntryListStruct elEntryList = { 0 };
	CBinaryParser Parser(cbBin, lpBin);

	Parser.GetDWORD(&elEntryList.EntryCount);
	Parser.GetDWORD(&elEntryList.Pad);

	if (elEntryList.EntryCount && elEntryList.EntryCount < _MaxEntriesLarge)
	{
		elEntryList.Entry = new EntryListEntryStruct[elEntryList.EntryCount];
		if (elEntryList.Entry)
		{
			memset(elEntryList.Entry, 0, sizeof(EntryListEntryStruct)* elEntryList.EntryCount);
			DWORD i = 0;
			for (i = 0; i < elEntryList.EntryCount; i++)
			{
				Parser.GetDWORD(&elEntryList.Entry[i].EntryLength);
				Parser.GetDWORD(&elEntryList.Entry[i].EntryLengthPad);
			}

			for (i = 0; i < elEntryList.EntryCount; i++)
			{
				size_t cbOffset = Parser.GetCurrentOffset();
				size_t cbRemainingBytes = Parser.RemainingBytes();
				cbRemainingBytes = min(elEntryList.Entry[i].EntryLength, cbRemainingBytes);
				elEntryList.Entry[i].EntryId = BinToEntryIdStruct(
					(ULONG)cbRemainingBytes,
					lpBin + cbOffset);
				Parser.Advance(cbRemainingBytes);
			}
		}
	}

	elEntryList.JunkDataSize = Parser.GetRemainingData(&elEntryList.JunkData);

	EntryListStruct* pelEntryList = new EntryListStruct;
	if (pelEntryList)
	{
		*pelEntryList = elEntryList;
	}

	return pelEntryList;
} // BinToEntryListStruct

void DeleteEntryListStruct(_In_ EntryListStruct* pelEntryList)
{
	if (!pelEntryList) return;
	if (pelEntryList->Entry)
	{
		DWORD i = 0;
		for (i = 0; i < pelEntryList->EntryCount; i++)
		{
			DeleteEntryIdStruct(pelEntryList->Entry[i].EntryId);
		}
	}
	delete[] pelEntryList->Entry;

	delete[] pelEntryList->JunkData;
	delete pelEntryList;
} // DeleteEntryListStruct

// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR EntryListStructToString(_In_ EntryListStruct* pelEntryList)
{
	if (!pelEntryList) return NULL;

	CString szEntryList;
	CString szTmp;

	szEntryList.FormatMessage(IDS_ENTRYLISTDATA,
		pelEntryList->EntryCount,
		pelEntryList->Pad);

	if (pelEntryList->Entry)
	{
		DWORD i = pelEntryList->EntryCount;
		for (i = 0; i < pelEntryList->EntryCount; i++)
		{
			szTmp.FormatMessage(IDS_ENTRYLISTENTRYID,
				i,
				pelEntryList->Entry[i].EntryLength,
				pelEntryList->Entry[i].EntryLengthPad);
			szEntryList += szTmp;
			LPWSTR szEntryId = EntryIdStructToString(pelEntryList->Entry[i].EntryId);
			szEntryList += szEntryId;
			delete[] szEntryId;
		}
	}

	szEntryList += JunkDataToString(pelEntryList->JunkDataSize, pelEntryList->JunkData);

	return CStringToLPWSTR(szEntryList);
} // EntryListStructToString

//////////////////////////////////////////////////////////////////////////
// End EntryListStruct
//////////////////////////////////////////////////////////////////////////