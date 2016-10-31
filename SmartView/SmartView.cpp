#include "stdafx.h"
#include "SmartView.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"
#include "MAPIFunctions.h"
#include "String.h"
#include "Guids.h"
#include "NamedPropCache.h"

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
#include "ExtendedFlags.h"
#include "AppointmentRecurrencePattern.h"
#include "RecurrencePattern.h"
#include "SIDBin.h"
#include "SDBin.h"
#include "XID.h"

wstring InterpretMVLongAsString(SLongArray myLongArray, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_ LPGUID lpguidNamedProp);

// Functions to parse PT_LONG/PT-I2 properties
_Check_return_ wstring RTimeToSzString(DWORD rTime, bool bLabel);
_Check_return_ wstring PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel);
// End: Functions to parse PT_LONG/PT-I2 properties

LPSMARTVIEWPARSER GetSmartViewParser(__ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp)
{
	switch (iStructType)
	{
	case IDS_STNOPARSING:
		return nullptr;
	case IDS_STTOMBSTONE:
		return new TombStone();
	case IDS_STPCL:
		return new PCL();
	case IDS_STVERBSTREAM:
		return new VerbStream();
	case IDS_STNICKNAMECACHE:
		return new NickNameCache();
	case IDS_STFOLDERUSERFIELDS:
		return new FolderUserFieldStream();
	case IDS_STRECIPIENTROWSTREAM:
		return new RecipientRowStream();
	case IDS_STWEBVIEWPERSISTSTREAM:
		return new WebViewPersistStream();
	case IDS_STFLATENTRYLIST:
		return new FlatEntryList();
	case IDS_STADDITIONALRENENTRYIDSEX:
		return new AdditionalRenEntryIDs();
	case IDS_STPROPERTYDEFINITIONSTREAM:
		return new PropertyDefinitionStream();
	case IDS_STSEARCHFOLDERDEFINITION:
		return new SearchFolderDefinition();
	case IDS_STENTRYLIST:
		return new EntryList();
	case IDS_STRULECONDITION:
	{
		auto parser = new RuleCondition();
		if (parser) parser->Init(false);
		return parser;
	}
	case IDS_STEXTENDEDRULECONDITION:
	{
		auto parser = new RuleCondition();
		if (parser) parser->Init(true);
		return parser;
	}
	case IDS_STRESTRICTION:
	{
		auto parser = new RestrictionStruct();
		if (parser) parser->Init(false, true);
		return parser;
	}
	case IDS_STPROPERTY:
		return new PropertyStruct();
	case IDS_STENTRYID:
		return new EntryIdStruct();
	case IDS_STGLOBALOBJECTID:
		return new GlobalObjectId();
	case IDS_STTASKASSIGNERS:
		return new TaskAssigners();
	case IDS_STCONVERSATIONINDEX:
		return new ConversationIndex();
	case IDS_STREPORTTAG:
		return new ReportTag();
	case IDS_STTIMEZONEDEFINITION:
		return new TimeZoneDefinition();
	case IDS_STTIMEZONE:
		return new TimeZone();
	case IDS_STEXTENDEDFOLDERFLAGS:
		return new ExtendedFlags();
	case IDS_STAPPOINTMENTRECURRENCEPATTERN:
		return new AppointmentRecurrencePattern();
	case IDS_STRECURRENCEPATTERN:
		return new RecurrencePattern();
	case IDS_STSID:
		return new SIDBin();
	case IDS_STSECURITYDESCRIPTOR:
	{
		auto parser = new SDBin();
		if (parser) parser->Init(lpMAPIProp, false);
		return parser;
	}
	case IDS_STFBSECURITYDESCRIPTOR:
	{
		auto parser = new SDBin();
		if (parser) parser->Init(lpMAPIProp, true);
		return parser;
	}
	case IDS_STXID:
		return new XID();
	}

	return nullptr;
}

_Check_return_ ULONG BuildFlagIndexFromTag(ULONG ulPropTag,
	ULONG ulPropNameID,
	_In_opt_z_ LPWSTR lpszPropNameString,
	_In_opt_ LPCGUID lpguidNamedProp)
{
	auto ulPropID = PROP_ID(ulPropTag);

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
}

_Check_return_ __ParsingTypeEnum FindSmartViewParserForProp(const ULONG ulPropTag, const ULONG ulPropNameID, _In_opt_ const LPCGUID lpguidNamedProp)
{
	auto ulIndex = BuildFlagIndexFromTag(ulPropTag, ulPropNameID, nullptr, lpguidNamedProp);
	auto bMV = (PROP_TYPE(ulPropTag) & MV_FLAG) == MV_FLAG;

	for (auto smartViewParser : SmartViewParserArray)
	{
		if (smartViewParser.ulIndex == ulIndex &&
			smartViewParser.bMV == bMV)
			return smartViewParser.iStructType;
	}

	return IDS_STNOPARSING;
}

_Check_return_ __ParsingTypeEnum FindSmartViewParserForProp(const ULONG ulPropTag, const ULONG ulPropNameID, _In_opt_ const LPCGUID lpguidNamedProp, bool bMVRow)
{
	auto ulLookupPropTag = ulPropTag;
	if (bMVRow) ulLookupPropTag |= MV_FLAG;

	return FindSmartViewParserForProp(ulLookupPropTag, ulPropNameID, lpguidNamedProp);
}

pair<__ParsingTypeEnum, wstring> InterpretPropSmartView2(_In_opt_ LPSPropValue lpProp, // required property value
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	bool bMVRow) // did the row come from a MV prop?
{
	wstring lpszSmartView;

	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD || !lpProp) return make_pair(IDS_STNOPARSING, L"");

	auto hRes = S_OK;
	auto iStructType = IDS_STNOPARSING;

	// Named Props
	LPMAPINAMEID* lppPropNames = nullptr;

	// If we weren't passed named property information and we need it, look it up
	// We check bIsAB here - some address book providers return garbage which will crash us
	if (!lpNameID &&
		lpMAPIProp && // if we have an object
		!bIsAB &&
		RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD && // and we're parsing named props
		(RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD || PROP_ID(lpProp->ulPropTag) >= 0x8000)) // and it's either a named prop or we're doing all props
	{
		SPropTagArray tag = { 0 };
		auto lpTag = &tag;
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
	LPGUID lpPropNameGUID = nullptr;

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
		lpszSmartView = InterpretNumberAsString(lpProp->Value, lpProp->ulPropTag, ulPropNameID, nullptr, lpPropNameGUID, true);
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
			LPSPropValue lpPropSubject = nullptr;

			WC_MAPI(HrGetOneProp(
				lpMAPIProp,
				PR_SUBJECT_W,
				&lpPropSubject));

			if (CheckStringProp(lpPropSubject, PT_UNICODE) && 0 == wcscmp(lpPropSubject->Value.lpszW, L"IPM.Configuration.Autocomplete"))
			{
				iStructType = IDS_STNICKNAMECACHE;
			}

			MAPIFreeBuffer(lpPropSubject);
		}

		if (iStructType)
		{
			lpszSmartView = InterpretBinaryAsString(lpProp->Value.bin, iStructType, lpMAPIProp);
		}

		break;
	case PT_MV_BINARY:
		iStructType = FindSmartViewParserForProp(lpProp->ulPropTag, ulPropNameID, lpPropNameGUID);
		if (iStructType)
		{
			lpszSmartView = InterpretMVBinaryAsString(lpProp->Value.MVbin, iStructType, lpMAPIProp);
		}

		break;
	}

	if (lppPropNames) MAPIFreeBuffer(lppPropNames);

	return make_pair(iStructType, lpszSmartView);
}

wstring InterpretPropSmartView(_In_opt_ LPSPropValue lpProp, // required property value
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	bool bMVRow) // did the row come from a MV prop?
{
	auto smartview = InterpretPropSmartView2(lpProp, lpMAPIProp, lpNameID, lpMappingSignature, bIsAB, bMVRow);
	return smartview.second;
}

wstring InterpretMVBinaryAsString(SBinaryArray myBinArray, __ParsingTypeEnum  iStructType, _In_opt_ LPMAPIPROP lpMAPIProp)
{
	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return L"";

	wstring szResult;

	for (ULONG ulRow = 0; ulRow < myBinArray.cValues; ulRow++)
	{
		if (ulRow != 0)
		{
			szResult += L"\r\n\r\n"; // STRING_OK
		}

		szResult += formatmessage(IDS_MVROWBIN, ulRow);
		szResult += InterpretBinaryAsString(myBinArray.lpbin[ulRow], iStructType, lpMAPIProp);
	}

	return szResult;
}

wstring InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag)
{
	_PV pV = { 0 };
	pV.ul = ulVal;
	return InterpretNumberAsString(pV, ulPropTag, NULL, nullptr, nullptr, false);
}

wstring InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp)
{
	_PV pV = { 0 };
	pV.ul = ulVal;
	return InterpretNumberAsString(pV, PT_LONG, ulPropNameID, nullptr, lpguidNamedProp, false);
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
	auto iParser = FindSmartViewParserForProp(ulPropTag, ulPropNameID, lpguidNamedProp);
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
			lpszResultString += InterpretFlags(ulPropID, pV.ul);

			if (bLabel && !lpszResultString.empty())
			{
				lpszResultString = formatmessage(IDS_FLAGS_PREFIX) + lpszResultString;
			}
		}

		break;
	}

	return lpszResultString;
}

wstring InterpretMVLongAsString(SLongArray myLongArray, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_ LPGUID lpguidNamedProp)
{
	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return L"";

	wstring szResult;
	wstring szSmartView;
	auto bHasData = false;

	for (ULONG ulRow = 0; ulRow < myLongArray.cValues; ulRow++)
	{
		_PV pV = { 0 };
		pV.ul = myLongArray.lpl[ulRow];
		szSmartView = InterpretNumberAsString(pV, CHANGE_PROP_TYPE(ulPropTag, PT_LONG), ulPropNameID, nullptr, lpguidNamedProp, true);
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

wstring InterpretBinaryAsString(SBinary myBin, __ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp)
{
	if (!RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD) return L"";
	auto szResultString = AddInSmartView(iStructType, myBin.cb, myBin.lpb);
	if (!szResultString.empty())
	{
		return szResultString;
	}

	auto svp = GetSmartViewParser(iStructType, lpMAPIProp);
	if (svp)
	{
		svp->Init(myBin.cb, myBin.lpb);
		szResultString = svp->ToString();
		delete svp;
		return szResultString;
	}

	// These parsers have some special casing
	switch (iStructType)
	{
	case IDS_STDECODEENTRYID:
		szResultString = DecodeID(myBin.cb, myBin.lpb);
		break;
	case IDS_STENCODEENTRYID:
		szResultString = EncodeID(myBin.cb, reinterpret_cast<LPENTRYID>(myBin.lpb));
		break;
	}

	return szResultString;
}

_Check_return_ wstring RTimeToString(DWORD rTime)
{
	wstring rTimeString;
	wstring rTimeAltString;
	FILETIME fTime = { 0 };
	LARGE_INTEGER liNumSec = { 0 };
	liNumSec.LowPart = rTime;
	// Resolution of RTime is in minutes, FILETIME is in 100 nanosecond intervals
	// Scale between the two is 10000000*60
	liNumSec.QuadPart = liNumSec.QuadPart * 10000000 * 60;
	fTime.dwLowDateTime = liNumSec.LowPart;
	fTime.dwHighDateTime = liNumSec.HighPart;
	FileTimeToString(&fTime, rTimeString, rTimeAltString);
	return rTimeString;
}

_Check_return_ wstring RTimeToSzString(DWORD rTime, bool bLabel)
{
	wstring szRTime;
	if (bLabel)
	{
		szRTime = L"RTime: "; // STRING_OK
	}

	szRTime += RTimeToString(rTime);
	return szRTime;
}

_Check_return_ wstring PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel)
{
	if (bLabel)
	{
		return formatmessage(IDS_PTI8FORMATLABEL, liI8.LowPart, liI8.HighPart);
	}

	return formatmessage(IDS_PTI8FORMAT, liI8.LowPart, liI8.HighPart);
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
}

_Check_return_ ULONGLONG UllGetIdGlobcnt(ID id)
{
	ULONGLONG ul = 0;
	for (auto i = 0; i < cbGlobcnt; ++i)
	{
		ul <<= 8;
		ul += id.globcnt[i];
	}

	return ul;
}

_Check_return_ wstring FidMidToSzString(LONGLONG llID, bool bLabel)
{
	auto pid = reinterpret_cast<ID*>(&llID);
	if (bLabel)
	{
		return formatmessage(IDS_FIDMIDFORMATLABEL, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
	}

	return formatmessage(IDS_FIDMIDFORMAT, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
}
