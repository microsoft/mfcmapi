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
#include "ExtendedFlags.h"
#include "AppointmentRecurrencePattern.h"
#include "RecurrencePattern.h"

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
	case IDS_STEXTENDEDFOLDERFLAGS:
		return new ExtendedFlags(cbBin, lpBin);
		break;
	case IDS_STAPPOINTMENTRECURRENCEPATTERN:
		return new AppointmentRecurrencePattern(cbBin, lpBin);
		break;
	case IDS_STRECURRENCEPATTERN:
		return new RecurrencePattern(cbBin, lpBin);
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
}

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
