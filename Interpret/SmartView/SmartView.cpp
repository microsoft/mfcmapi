#include <StdAfx.h>
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/String.h>
#include <Interpret/Guids.h>
#include <MAPI/Cache/NamedPropCache.h>

#include <Interpret/SmartView/SmartViewParser.h>
#include <Interpret/SmartView/PCL.h>
#include <Interpret/SmartView/TombStone.h>
#include <Interpret/SmartView/VerbStream.h>
#include <Interpret/SmartView/NickNameCache.h>
#include <Interpret/SmartView/FolderUserFieldStream.h>
#include <Interpret/SmartView/RecipientRowStream.h>
#include <Interpret/SmartView/WebViewPersistStream.h>
#include <Interpret/SmartView/FlatEntryList.h>
#include <Interpret/SmartView/AdditionalRenEntryIDs.h>
#include <Interpret/SmartView/PropertyDefinitionStream.h>
#include <Interpret/SmartView/SearchFolderDefinition.h>
#include <Interpret/SmartView/EntryList.h>
#include <Interpret/SmartView/RuleCondition.h>
#include <Interpret/SmartView/RestrictionStruct.h>
#include <Interpret/SmartView/PropertyStruct.h>
#include <Interpret/SmartView/EntryIdStruct.h>
#include <Interpret/SmartView/GlobalObjectId.h>
#include <Interpret/SmartView/TaskAssigners.h>
#include <Interpret/SmartView/ConversationIndex.h>
#include <Interpret/SmartView/ReportTag.h>
#include <Interpret/SmartView/TimeZoneDefinition.h>
#include <Interpret/SmartView/TimeZone.h>
#include <Interpret/SmartView/ExtendedFlags.h>
#include <Interpret/SmartView/AppointmentRecurrencePattern.h>
#include <Interpret/SmartView/RecurrencePattern.h>
#include <Interpret/SmartView/SIDBin.h>
#include <Interpret/SmartView/SDBin.h>
#include <Interpret/SmartView/XID.h>

namespace smartview
{
	std::wstring InterpretMVLongAsString(
		SLongArray myLongArray,
		ULONG ulPropTag,
		ULONG ulPropNameID,
		_In_opt_ LPGUID lpguidNamedProp);

	// Functions to parse PT_LONG/PT-I2 properties
	_Check_return_ std::wstring RTimeToSzString(DWORD rTime, bool bLabel);
	_Check_return_ std::wstring PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel);
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
			auto parser = new (std::nothrow) RuleCondition();
			if (parser) parser->Init(false);
			return parser;
		}
		case IDS_STEXTENDEDRULECONDITION:
		{
			auto parser = new (std::nothrow) RuleCondition();
			if (parser) parser->Init(true);
			return parser;
		}
		case IDS_STRESTRICTION:
		{
			auto parser = new (std::nothrow) RestrictionStruct();
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
			auto parser = new (std::nothrow) SDBin();
			if (parser) parser->Init(lpMAPIProp, false);
			return parser;
		}
		case IDS_STFBSECURITYDESCRIPTOR:
		{
			auto parser = new (std::nothrow) SDBin();
			if (parser) parser->Init(lpMAPIProp, true);
			return parser;
		}
		case IDS_STXID:
			return new XID();
		}

		return nullptr;
	}

	_Check_return_ ULONG BuildFlagIndexFromTag(
		ULONG ulPropTag,
		ULONG ulPropNameID,
		_In_opt_z_ LPCWSTR lpszPropNameString,
		_In_opt_ LPCGUID lpguidNamedProp)
	{
		const auto ulPropID = PROP_ID(ulPropTag);

		// Non-zero less than 0x8000 is a regular prop, we use the ID as the index
		if (ulPropID && ulPropID < 0x8000) return ulPropID;

		// Else we build our index from the guid and named property ID
		// In the future, we can look at using lpszPropNameString for MNID_STRING named properties
		if (lpguidNamedProp && (ulPropNameID || lpszPropNameString))
		{
			ULONG ulGuid = NULL;
			if (*lpguidNamedProp == guid::PSETID_Meeting)
				ulGuid = guidPSETID_Meeting;
			else if (*lpguidNamedProp == guid::PSETID_Address)
				ulGuid = guidPSETID_Address;
			else if (*lpguidNamedProp == guid::PSETID_Task)
				ulGuid = guidPSETID_Task;
			else if (*lpguidNamedProp == guid::PSETID_Appointment)
				ulGuid = guidPSETID_Appointment;
			else if (*lpguidNamedProp == guid::PSETID_Common)
				ulGuid = guidPSETID_Common;
			else if (*lpguidNamedProp == guid::PSETID_Log)
				ulGuid = guidPSETID_Log;
			else if (*lpguidNamedProp == guid::PSETID_PostRss)
				ulGuid = guidPSETID_PostRss;
			else if (*lpguidNamedProp == guid::PSETID_Sharing)
				ulGuid = guidPSETID_Sharing;
			else if (*lpguidNamedProp == guid::PSETID_Note)
				ulGuid = guidPSETID_Note;

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

	_Check_return_ __ParsingTypeEnum
	FindSmartViewParserForProp(const ULONG ulPropTag, const ULONG ulPropNameID, _In_opt_ const LPCGUID lpguidNamedProp)
	{
		const auto ulIndex = BuildFlagIndexFromTag(ulPropTag, ulPropNameID, nullptr, lpguidNamedProp);
		const auto bMV = (PROP_TYPE(ulPropTag) & MV_FLAG) == MV_FLAG;

		for (const auto& smartViewParser : SmartViewParserArray)
		{
			if (smartViewParser.ulIndex == ulIndex && smartViewParser.bMV == bMV) return smartViewParser.iStructType;
		}

		return IDS_STNOPARSING;
	}

	_Check_return_ __ParsingTypeEnum FindSmartViewParserForProp(
		const ULONG ulPropTag,
		const ULONG ulPropNameID,
		_In_opt_ const LPCGUID lpguidNamedProp,
		bool bMVRow)
	{
		auto ulLookupPropTag = ulPropTag;
		if (bMVRow) ulLookupPropTag |= MV_FLAG;

		return FindSmartViewParserForProp(ulLookupPropTag, ulPropNameID, lpguidNamedProp);
	}

	std::pair<__ParsingTypeEnum, std::wstring> InterpretPropSmartView2(
		_In_opt_ const _SPropValue* lpProp, // required property value
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
		bool bMVRow) // did the row come from a MV prop?
	{
		std::wstring lpszSmartView;

		if (!registry::RegKeys[registry::regkeyDO_SMART_VIEW].ulCurDWORD || !lpProp)
			return std::make_pair(IDS_STNOPARSING, L"");

		auto iStructType = IDS_STNOPARSING;

		// Named Props
		LPMAPINAMEID* lppPropNames = nullptr;

		// If we weren't passed named property information and we need it, look it up
		// We check bIsAB here - some address book providers return garbage which will crash us
		if (!lpNameID && lpMAPIProp && // if we have an object
			!bIsAB &&
			registry::RegKeys[registry::regkeyPARSED_NAMED_PROPS].ulCurDWORD && // and we're parsing named props
			(registry::RegKeys[registry::regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD ||
			 PROP_ID(lpProp->ulPropTag) >= 0x8000)) // and it's either a named prop or we're doing all props
		{
			SPropTagArray tag = {0};
			auto lpTag = &tag;
			ULONG ulPropNames = 0;
			tag.cValues = 1;
			tag.aulPropTag[0] = lpProp->ulPropTag;

			WC_H_GETPROPS_S(cache::GetNamesFromIDs(
				lpMAPIProp, lpMappingSignature, &lpTag, nullptr, NULL, &ulPropNames, &lppPropNames));
			if (ulPropNames == 1 && lppPropNames && lppPropNames[0])
			{
				lpNameID = lppPropNames[0];
			}
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
			lpszSmartView =
				InterpretNumberAsString(lpProp->Value, lpProp->ulPropTag, ulPropNameID, nullptr, lpPropNameGUID, true);
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

				WC_MAPI_S(HrGetOneProp(lpMAPIProp, PR_SUBJECT_W, &lpPropSubject));

				if (mapi::CheckStringProp(lpPropSubject, PT_UNICODE) &&
					0 == wcscmp(lpPropSubject->Value.lpszW, L"IPM.Configuration.Autocomplete"))
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

	std::wstring InterpretPropSmartView(
		_In_opt_ LPSPropValue lpProp, // required property value
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
		bool bMVRow) // did the row come from a MV prop?
	{
		auto smartview = InterpretPropSmartView2(lpProp, lpMAPIProp, lpNameID, lpMappingSignature, bIsAB, bMVRow);
		return smartview.second;
	}

	std::wstring
	InterpretMVBinaryAsString(SBinaryArray myBinArray, __ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp)
	{
		if (!registry::RegKeys[registry::regkeyDO_SMART_VIEW].ulCurDWORD) return L"";

		std::wstring szResult;

		for (ULONG ulRow = 0; ulRow < myBinArray.cValues; ulRow++)
		{
			if (ulRow != 0)
			{
				szResult += L"\r\n\r\n"; // STRING_OK
			}

			szResult += strings::formatmessage(IDS_MVROWBIN, ulRow);
			szResult += InterpretBinaryAsString(myBinArray.lpbin[ulRow], iStructType, lpMAPIProp);
		}

		return szResult;
	}

	std::wstring InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag)
	{
		_PV pV = {0};
		pV.ul = ulVal;
		return InterpretNumberAsString(pV, ulPropTag, NULL, nullptr, nullptr, false);
	}

	std::wstring InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp)
	{
		_PV pV = {0};
		pV.ul = ulVal;
		return InterpretNumberAsString(pV, PT_LONG, ulPropNameID, nullptr, lpguidNamedProp, false);
	}

	// Interprets a PT_LONG, PT_I2. or PT_I8 found in lpProp and returns a string
	// Will not return a string if the lpProp is not a PT_LONG/PT_I2/PT_I8 or we don't recognize the property
	// Will use named property details to look up named property flags
	std::wstring InterpretNumberAsString(
		_PV pV,
		ULONG ulPropTag,
		ULONG ulPropNameID,
		_In_opt_z_ LPWSTR lpszPropNameString,
		_In_opt_ LPCGUID lpguidNamedProp,
		bool bLabel)
	{
		std::wstring lpszResultString;
		if (!ulPropTag) return L"";

		if (PROP_TYPE(ulPropTag) != PT_LONG && //
			PROP_TYPE(ulPropTag) != PT_I2 && //
			PROP_TYPE(ulPropTag) != PT_I8)
			return L"";

		ULONG ulPropID = NULL;
		const auto iParser = FindSmartViewParserForProp(ulPropTag, ulPropNameID, lpguidNamedProp);
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
				lpszResultString += interpretprop::InterpretFlags(ulPropID, pV.ul);

				if (bLabel && !lpszResultString.empty())
				{
					lpszResultString = strings::formatmessage(IDS_FLAGS_PREFIX) + lpszResultString;
				}
			}

			break;
		}

		return lpszResultString;
	}

	std::wstring InterpretMVLongAsString(
		SLongArray myLongArray,
		ULONG ulPropTag,
		ULONG ulPropNameID,
		_In_opt_ LPGUID lpguidNamedProp)
	{
		if (!registry::RegKeys[registry::regkeyDO_SMART_VIEW].ulCurDWORD) return L"";

		std::wstring szResult;
		std::wstring szSmartView;
		auto bHasData = false;

		for (ULONG ulRow = 0; ulRow < myLongArray.cValues; ulRow++)
		{
			_PV pV = {0};
			pV.ul = myLongArray.lpl[ulRow];
			szSmartView = InterpretNumberAsString(
				pV, CHANGE_PROP_TYPE(ulPropTag, PT_LONG), ulPropNameID, nullptr, lpguidNamedProp, true);
			if (!szSmartView.empty())
			{
				if (bHasData)
				{
					szResult += L"\r\n"; // STRING_OK
				}

				bHasData = true;
				szResult += strings::formatmessage(IDS_MVROWLONG, ulRow, szSmartView.c_str());
			}
		}

		return szResult;
	}

	std::wstring InterpretBinaryAsString(SBinary myBin, __ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp)
	{
		if (!registry::RegKeys[registry::regkeyDO_SMART_VIEW].ulCurDWORD) return L"";
		auto szResultString = addin::AddInSmartView(iStructType, myBin.cb, myBin.lpb);
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
			szResultString = mapi::DecodeID(myBin.cb, myBin.lpb);
			break;
		case IDS_STENCODEENTRYID:
			szResultString = mapi::EncodeID(myBin.cb, reinterpret_cast<LPENTRYID>(myBin.lpb));
			break;
		}

		return szResultString;
	}

	_Check_return_ std::wstring RTimeToString(DWORD rTime)
	{
		std::wstring rTimeString;
		std::wstring rTimeAltString;
		FILETIME fTime = {0};
		LARGE_INTEGER liNumSec = {0};
		liNumSec.LowPart = rTime;
		// Resolution of RTime is in minutes, FILETIME is in 100 nanosecond intervals
		// Scale between the two is 10000000*60
		liNumSec.QuadPart = liNumSec.QuadPart * 10000000 * 60;
		fTime.dwLowDateTime = liNumSec.LowPart;
		fTime.dwHighDateTime = liNumSec.HighPart;
		strings::FileTimeToString(fTime, rTimeString, rTimeAltString);
		return rTimeString;
	}

	_Check_return_ std::wstring RTimeToSzString(DWORD rTime, bool bLabel)
	{
		std::wstring szRTime;
		if (bLabel)
		{
			szRTime = L"RTime: "; // STRING_OK
		}

		szRTime += RTimeToString(rTime);
		return szRTime;
	}

	_Check_return_ std::wstring PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel)
	{
		if (bLabel)
		{
			return strings::formatmessage(IDS_PTI8FORMATLABEL, liI8.LowPart, liI8.HighPart);
		}

		return strings::formatmessage(IDS_PTI8FORMAT, liI8.LowPart, liI8.HighPart);
	}

	typedef WORD REPLID;
#define cbGlobcnt 6

	struct ID
	{
		REPLID replid;
		BYTE globcnt[cbGlobcnt];
	};

	_Check_return_ WORD WGetReplId(ID id) { return id.replid; }

	_Check_return_ ULONGLONG UllGetIdGlobcnt(ID id)
	{
		ULONGLONG ul = 0;
		for (auto i : id.globcnt)
		{
			ul <<= 8;
			ul += i;
		}

		return ul;
	}

	_Check_return_ std::wstring FidMidToSzString(LONGLONG llID, bool bLabel)
	{
		const auto pid = reinterpret_cast<ID*>(&llID);
		if (bLabel)
		{
			return strings::formatmessage(IDS_FIDMIDFORMATLABEL, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
		}

		return strings::formatmessage(IDS_FIDMIDFORMAT, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
	}
} // namespace smartview