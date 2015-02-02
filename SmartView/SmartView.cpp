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
#include "EntryList.h"
#include "RuleCondition.h"
#include "RestrictionStruct.h"
#include "PropertyStruct.h"
#include "EntryIdStruct.h"
#include "GlobalObjectId.h"
#include "TaskAssigners.h"
#include "ConversationIndex.h"
#include "ReportTag.h"
#include "TimeZoneDefinition.h"
#include "TimeZone.h"

#define _MaxBytes 0xFFFF
#define _MaxDepth 50
#define _MaxEID 500
#define _MaxEntriesSmall 500
#define _MaxEntriesLarge 1000
#define _MaxEntriesExtraLarge 1500
#define _MaxEntriesEnormous 10000

wstring InterpretMVLongAsString(SLongArray myLongArray, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_ LPGUID lpguidNamedProp);

// Functions to parse PT_LONG/PT-I2 properties

_Check_return_ CString RTimeToString(DWORD rTime);
_Check_return_ wstring RTimeToSzString(DWORD rTime, bool bLabel);
_Check_return_ wstring PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel);

// End: Functions to parse PT_LONG/PT-I2 properties

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
	__ParsingTypeEnum iStructType;
	LPBINTOSTRUCT fBinToStruct;
	LPDELETESTRUCT fDeleteStruct;
	LPSTRUCTTOSTRING fStructToString;
};
#define MAKE_SV_ENTRY(_fIDSType, _fType) {(_fIDSType), (LPBINTOSTRUCT) BinTo##_fType, (LPDELETESTRUCT) Delete##_fType, (LPSTRUCTTOSTRING) _fType##ToString},
SMART_VIEW_PARSERS_ENTRY g_SmartViewParsers[] = {
	{ IDS_STNOPARSING, NULL, NULL, NULL },
	// MAKE_SV_ENTRY(IDS_STSECURITYDESCRIPTOR, NULL}
	MAKE_SV_ENTRY(IDS_STEXTENDEDFOLDERFLAGS, ExtendedFlagsStruct)
	MAKE_SV_ENTRY(IDS_STAPPOINTMENTRECURRENCEPATTERN, AppointmentRecurrencePatternStruct)
	MAKE_SV_ENTRY(IDS_STRECURRENCEPATTERN, RecurrencePatternStruct)
	// MAKE_SV_ENTRY(IDS_STSID, SIDStruct)
	// MAKE_SV_ENTRY(IDS_STDECODEENTRYID)
	// MAKE_SV_ENTRY(IDS_STENCODEENTRYID)
};
ULONG g_cSmartViewParsers = _countof(g_SmartViewParsers);

LPSMARTVIEWPARSER GetSmartViewParser(__ParsingTypeEnum iStructType, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
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
	case IDS_STENTRYLIST:
		return new EntryList(cbBin, lpBin);
		break;
	case IDS_STRULECONDITION:
		return new RuleCondition(cbBin, lpBin, false);
		break;
	case IDS_STEXTENDEDRULECONDITION:
		return new RuleCondition(cbBin, lpBin, true);
		break;
	case IDS_STRESTRICTION:
		return new RestrictionStruct(cbBin, lpBin);
		break;
	case IDS_STPROPERTY:
		return new PropertyStruct(cbBin, lpBin);
		break;
	case IDS_STENTRYID:
		return new EntryIdStruct(cbBin, lpBin);
		break;
	case IDS_STGLOBALOBJECTID:
		return new GlobalObjectId(cbBin, lpBin);
		break;
	case IDS_STTASKASSIGNERS:
		return new TaskAssigners(cbBin, lpBin);
		break;
	case IDS_STCONVERSATIONINDEX:
		return new ConversationIndex(cbBin, lpBin);
		break;
	case IDS_STREPORTTAG:
		return new ReportTag(cbBin, lpBin);
		break;
	case IDS_STTIMEZONEDEFINITION:
		return new TimeZoneDefinition(cbBin, lpBin);
		break;
	case IDS_STTIMEZONE:
		return new TimeZone(cbBin, lpBin);
		break;
	}

	return NULL;
}

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

_Check_return_ __ParsingTypeEnum FindSmartViewParserForProp(const ULONG ulPropTag, const ULONG ulPropNameID, _In_opt_ const LPCGUID lpguidNamedProp)
{
	ULONG ulCurEntry = 0;
	ULONG ulIndex = BuildFlagIndexFromTag(ulPropTag, ulPropNameID, NULL, lpguidNamedProp);
	bool bMV = (PROP_TYPE(ulPropTag) & MV_FLAG) == MV_FLAG;

	while (ulCurEntry < ulSmartViewParserArray)
	{
		if (SmartViewParserArray[ulCurEntry].ulIndex == ulIndex &&
			SmartViewParserArray[ulCurEntry].bMV == bMV)
			return SmartViewParserArray[ulCurEntry].iStructType;
		ulCurEntry++;
	}

	return IDS_STNOPARSING;
}

_Check_return_ __ParsingTypeEnum FindSmartViewParserForProp(const ULONG ulPropTag, const ULONG ulPropNameID, _In_opt_ const LPCGUID lpguidNamedProp, bool bMVRow)
{
	ULONG ulLookupPropTag = ulPropTag;
	if (bMVRow) ulLookupPropTag |= MV_FLAG;

	return FindSmartViewParserForProp(ulLookupPropTag, ulPropNameID, lpguidNamedProp);
}

wstring InterpretPropSmartView(_In_ LPSPropValue lpProp, // required property value
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	bool bMVRow) // did the row come from a MV prop?
{
	wstring lpszSmartView;

	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return L"";
	if (!lpProp) return L"";

	HRESULT hRes = S_OK;
	__ParsingTypeEnum iStructType = IDS_STNOPARSING;

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
		SPropTagArray tag = { 0 };
		LPSPropTagArray lpTag = &tag;
		ULONG ulPropNames = 0;
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

	ULONG ulLookupPropTag = NULL;
	switch (PROP_TYPE(lpProp->ulPropTag))
	{
	case PT_LONG:
	case PT_I2:
	case PT_I8:
		lpszSmartView = InterpretNumberAsString(lpProp->Value, lpProp->ulPropTag, ulPropNameID, NULL, lpPropNameGUID, true);
		break;
	case PT_MV_LONG:
		lpszSmartView = InterpretMVLongAsString(lpProp->Value.MVl, lpProp->ulPropTag, ulPropNameID, lpPropNameGUID);
		break;
	case PT_BINARY:
		ulLookupPropTag = lpProp->ulPropTag;
		if (bMVRow) ulLookupPropTag |= MV_FLAG;

		iStructType = FindSmartViewParserForProp(ulLookupPropTag, ulPropNameID, lpPropNameGUID);
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
			lpszSmartView = InterpretBinaryAsString(lpProp->Value.bin, iStructType, lpMAPIProp, lpProp->ulPropTag);
		}

		break;
	case PT_MV_BINARY:
		iStructType = FindSmartViewParserForProp(lpProp->ulPropTag, ulPropNameID, lpPropNameGUID);
		if (iStructType)
		{
			lpszSmartView = InterpretMVBinaryAsString(lpProp->Value.MVbin, iStructType, lpMAPIProp, lpProp->ulPropTag);
		}

		break;
	}
	MAPIFreeBuffer(lppPropNames);

	return lpszSmartView;
}

wstring InterpretMVBinaryAsString(SBinaryArray myBinArray, __ParsingTypeEnum  iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
{
	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return L"";

	ULONG ulRow = 0;
	wstring szResult;

	for (ulRow = 0; ulRow < myBinArray.cValues; ulRow++)
	{
		if (ulRow != 0)
		{
			szResult += L"\r\n\r\n"; // STRING_OK
		}

		szResult += formatmessage(IDS_MVROWBIN, ulRow);
		szResult += InterpretBinaryAsString(myBinArray.lpbin[ulRow], iStructType, lpMAPIProp, ulPropTag);
	}

	return szResult;
}

wstring InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag)
{
	_PV pV = { 0 };
	pV.ul = ulVal;
	return InterpretNumberAsString(pV, ulPropTag, NULL, NULL, NULL, false);
}

wstring InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp)
{
	_PV pV = { 0 };
	pV.ul = ulVal;
	return InterpretNumberAsString(pV, PT_LONG, ulPropNameID, NULL, lpguidNamedProp, false);
}

// Interprets a PT_LONG, PT_I2. or PT_I8 found in lpProp and returns a string
// Will not return a string if the lpProp is not a PT_LONG/PT_I2/PT_I8 or we don't recognize the property
// Will use named property details to look up named property flags
wstring InterpretNumberAsString(_PV pV, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_z_ LPWSTR lpszPropNameString, _In_opt_ LPCGUID lpguidNamedProp, bool bLabel)
{
	wstring lpszResultString;
	if (!ulPropTag) return L"";

	if (PROP_TYPE(ulPropTag) != PT_LONG &&
		PROP_TYPE(ulPropTag) != PT_I2 &&
		PROP_TYPE(ulPropTag) != PT_I8) return L"";

	ULONG ulPropID = NULL;
	__ParsingTypeEnum iParser = FindSmartViewParserForProp(ulPropTag, ulPropNameID, lpguidNamedProp);
	switch (iParser)
	{
	case IDS_STLONGRTIME:
		lpszResultString = RTimeToSzString(pV.ul, bLabel);
		break;
	case IDS_STPTI8:
		lpszResultString = PTI8ToSzString(pV.li, bLabel);
		break;
	case IDS_STSFIDMID:
		lpszResultString = FidMidToSzString(pV.li.QuadPart, bLabel);
		break;
		// insert future parsers here
	default:
		ulPropID = BuildFlagIndexFromTag(ulPropTag, ulPropNameID, lpszPropNameString, lpguidNamedProp);
		if (ulPropID)
		{
			CString szPrefix;
			HRESULT hRes = S_OK;
			if (bLabel)
			{
				EC_B(szPrefix.LoadString(IDS_FLAGS_PREFIX));
			}

			LPTSTR szTmp = NULL;
			InterpretFlags(ulPropID, pV.ul, szPrefix, &szTmp);
			lpszResultString = LPTSTRToWstring(szTmp);
			delete[] szTmp;
		}

		break;
	}

	return lpszResultString;
}

wstring InterpretMVLongAsString(SLongArray myLongArray, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_ LPGUID lpguidNamedProp)
{
	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return L"";

	ULONG ulRow = 0;
	wstring szResult;
	wstring szSmartView;
	bool bHasData = false;

	for (ulRow = 0; ulRow < myLongArray.cValues; ulRow++)
	{
		_PV pV = { 0 };
		pV.ul = myLongArray.lpl[ulRow];
		szSmartView = InterpretNumberAsString(pV, CHANGE_PROP_TYPE(ulPropTag, PT_LONG), ulPropNameID, NULL, lpguidNamedProp, true);
		if (!szSmartView.empty())
		{
			if (bHasData)
			{
				szResult += L"\r\n"; // STRING_OK
			}

			bHasData = true;
			szResult += formatmessage(IDS_MVROWLONG,
				ulRow,
				szSmartView.c_str());
		}
	}

	return szResult;
}

wstring InterpretBinaryAsString(SBinary myBin, __ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
{
	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return L"";
	wstring szResultString = AddInSmartView(iStructType, myBin.cb, myBin.lpb);
	if (!szResultString.empty())
	{
		return szResultString;
	}

	ULONG i = 0;
	for (i = 0; i < g_cSmartViewParsers; i++)
	{
		if (iStructType == g_SmartViewParsers[i].iStructType)
		{
			if (g_SmartViewParsers[i].fBinToStruct && g_SmartViewParsers[i].fStructToString && g_SmartViewParsers[i].fDeleteStruct)
			{
				LPVOID pStruct = g_SmartViewParsers[i].fBinToStruct(myBin.cb, myBin.lpb);
				if (pStruct)
				{
					LPWSTR szTmp = g_SmartViewParsers[i].fStructToString(pStruct);

					if (szTmp)
					{
						szResultString = szTmp;
						delete[] szTmp;
					}

					g_SmartViewParsers[i].fDeleteStruct(pStruct);
					pStruct = NULL;
				}
			}

			return szResultString;
		}
	}

	LPSMARTVIEWPARSER svp = GetSmartViewParser(iStructType, myBin.cb, myBin.lpb);
	if (svp)
	{
		szResultString = svp->ToString();
		delete svp;
		return szResultString;
	}

	// These parsers have some special casing
	LPWSTR szTmp = NULL;

	switch (iStructType)
	{
	case IDS_STSECURITYDESCRIPTOR:
		SDBinToString(myBin, lpMAPIProp, ulPropTag, &szTmp);
		break;
	case IDS_STSID:
		SIDBinToString(myBin, &szTmp);
		break;
	case IDS_STDECODEENTRYID:
		szTmp = DecodeID(myBin.cb, myBin.lpb);
		break;
	case IDS_STENCODEENTRYID:
		szTmp = EncodeID(myBin.cb, (LPENTRYID)myBin.lpb);
		break;
	}

	if (szTmp)
	{
		szResultString = szTmp;
		delete[] szTmp;
	}

	return szResultString;
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

_Check_return_ wstring RTimeToSzString(DWORD rTime, bool bLabel)
{
	wstring szRTime;
	if (bLabel)
	{
		szRTime = L"RTime: "; // STRING_OK
	}

	LPWSTR szTmp = CStringToLPWSTR(RTimeToString(rTime));
	szRTime += szTmp;
	delete[] szTmp;
	return szRTime;
}

_Check_return_ wstring PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel)
{
	if (bLabel)
	{
		return formatmessage(IDS_PTI8FORMATLABEL, liI8.LowPart, liI8.HighPart);
	}
	else
	{
		return formatmessage(IDS_PTI8FORMAT, liI8.LowPart, liI8.HighPart);
	}
}

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

_Check_return_ wstring FidMidToSzString(LONGLONG llID, bool bLabel)
{
	ID* pid = (ID*)&llID;
	if (bLabel)
	{
		return formatmessage(IDS_FIDMIDFORMATLABEL, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
	}
	else
	{
		return formatmessage(IDS_FIDMIDFORMAT, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
	}
}

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
				LPWSTR szFlags = wstringToLPWSTR(InterpretNumberAsStringNamedProp(parpPattern->ExceptionInfo[i].MeetingType, dispidApptStateFlags, (LPGUID)&PSETID_Appointment));
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
				LPWSTR szFlags = wstringToLPWSTR(InterpretNumberAsStringNamedProp(parpPattern->ExceptionInfo[i].BusyStatus, dispidBusyStatus, (LPGUID)&PSETID_Appointment));
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
				LPWSTR szFlags = wstringToLPWSTR(InterpretNumberAsStringNamedProp(parpPattern->ExtendedException[i].ChangeHighlight.ChangeHighlightValue, dispidChangeHighlight, (LPGUID)&PSETID_Appointment));
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
		myBin.cb >= sizeof(SID) - sizeof(DWORD) + sizeof(DWORD)* ((PISID)SidStart)->SubAuthorityCount &&
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
				szFlags = wstringToLPTSTR(GUIDToString(&pefExtendedFlags->pefExtendedFlags[i].Data.SearchFolderID));
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