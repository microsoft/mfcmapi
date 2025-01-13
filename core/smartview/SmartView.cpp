#include <core/stdafx.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>
#include <core/interpret/guid.h>
#include <core/mapi/cache/namedProps.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/registry.h>
#include <core/interpret/flags.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/error.h>

#include <core/smartview/PCL.h>
#include <core/smartview/TombStone.h>
#include <core/smartview/VerbStream.h>
#include <core/smartview/NickNameCache.h>
#include <core/smartview/FolderUserFieldStream.h>
#include <core/smartview/RecipientRowStream.h>
#include <core/smartview/WebViewPersistStream.h>
#include <core/smartview/FlatEntryList.h>
#include <core/smartview/AdditionalRenEntryIDs.h>
#include <core/smartview/PropertyDefinitionStream.h>
#include <core/smartview/SearchFolderDefinition.h>
#include <core/smartview/EntryList.h>
#include <core/smartview/RuleAction.h>
#include <core/smartview/RuleCondition.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/smartview/GlobalObjectId.h>
#include <core/smartview/TaskAssigners.h>
#include <core/smartview/ConversationIndex.h>
#include <core/smartview/ReportTag.h>
#include <core/smartview/TimeZoneDefinition.h>
#include <core/smartview/TimeZone.h>
#include <core/smartview/ExtendedFlags.h>
#include <core/smartview/AppointmentRecurrencePattern.h>
#include <core/smartview/RecurrencePattern.h>
#include <core/smartview/SD/SIDBin.h>
#include <core/smartview/SD/SDBin.h>
#include <core/smartview/XID.h>
#include <core/smartview/decodeEntryID.h>
#include <core/smartview/encodeEntryID.h>
#include <core/smartview/addinParser.h>
#include <core/smartview/swappedToDo.h>

namespace smartview
{
	std::shared_ptr<block> InterpretBinary(const SBinary myBin, parserType parser, _In_opt_ LPMAPIPROP lpMAPIProp)
	{
		if (!registry::doSmartView) emptySW();

		auto svp = GetSmartViewParser(parser, lpMAPIProp);
		if (svp)
		{
			svp->parse(std::make_shared<binaryParser>(myBin.cb, myBin.lpb), true);
			return svp;
		}

		return emptySW();
	}

	// Functions to parse PT_LONG/PT-I2 properties
	_Check_return_ std::wstring RTimeToSzString(DWORD rTime, bool bLabel);
	_Check_return_ std::wstring PTI8ToSzString(LARGE_INTEGER liI8, bool bLabel);
	// End: Functions to parse PT_LONG/PT-I2 properties

	std::shared_ptr<block> GetSmartViewParser(parserType type, _In_opt_ LPMAPIPROP lpMAPIProp)
	{
		switch (type)
		{
		case parserType::NOPARSING:
			return nullptr;
		case parserType::TOMBSTONE:
			return std::make_shared<TombStone>();
		case parserType::PCL:
			return std::make_shared<PCL>();
		case parserType::VERBSTREAM:
			return std::make_shared<VerbStream>();
		case parserType::NICKNAMECACHE:
			return std::make_shared<NickNameCache>();
		case parserType::DECODEENTRYID:
			return std::make_shared<decodeEntryID>();
		case parserType::ENCODEENTRYID:
			return std::make_shared<encodeEntryID>();
		case parserType::FOLDERUSERFIELDS:
			return std::make_shared<FolderUserFieldStream>();
		case parserType::RECIPIENTROWSTREAM:
			return std::make_shared<RecipientRowStream>();
		case parserType::WEBVIEWPERSISTSTREAM:
			return std::make_shared<WebViewPersistStream>();
		case parserType::FLATENTRYLIST:
			return std::make_shared<FlatEntryList>();
		case parserType::ADDITIONALRENENTRYIDSEX:
			return std::make_shared<AdditionalRenEntryIDs>();
		case parserType::PROPERTYDEFINITIONSTREAM:
			return std::make_shared<PropertyDefinitionStream>();
		case parserType::SEARCHFOLDERDEFINITION:
			return std::make_shared<SearchFolderDefinition>();
		case parserType::ENTRYLIST:
			return std::make_shared<EntryList>();
		case parserType::RULEACTION:
			return std::make_shared<RuleAction>(false);
		case parserType::EXTENDEDRULEACTION:
			return std::make_shared<RuleAction>(true);
		case parserType::RULECONDITION:
			return std::make_shared<RuleCondition>(false);
		case parserType::EXTENDEDRULECONDITION:
			return std::make_shared<RuleCondition>(true);
		case parserType::RESTRICTION:
			return std::make_shared<RestrictionStruct>(false, true);
		case parserType::PROPERTIES:
			return std::make_shared<PropertiesStruct>(_MaxEntriesSmall, false, false);
		case parserType::ENTRYID:
			return std::make_shared<EntryIdStruct>();
		case parserType::GLOBALOBJECTID:
			return std::make_shared<GlobalObjectId>();
		case parserType::TASKASSIGNERS:
			return std::make_shared<TaskAssigners>();
		case parserType::CONVERSATIONINDEX:
			return std::make_shared<ConversationIndex>();
		case parserType::REPORTTAG:
			return std::make_shared<ReportTag>();
		case parserType::TIMEZONEDEFINITION:
			return std::make_shared<TimeZoneDefinition>();
		case parserType::TIMEZONE:
			return std::make_shared<TimeZone>();
		case parserType::EXTENDEDFOLDERFLAGS:
			return std::make_shared<ExtendedFlags>();
		case parserType::APPOINTMENTRECURRENCEPATTERN:
			return std::make_shared<AppointmentRecurrencePattern>();
		case parserType::RECURRENCEPATTERN:
			return std::make_shared<RecurrencePattern>();
		case parserType::SID:
			return std::make_shared<SIDBin>();
		case parserType::SECURITYDESCRIPTOR:
			return std::make_shared<SDBin>(lpMAPIProp, false);
		case parserType::FBSECURITYDESCRIPTOR:
			return std::make_shared<SDBin>(lpMAPIProp, true);
		case parserType::XID:
			return std::make_shared<XID>();
		case parserType::SWAPPEDTODO:
			return std::make_shared<swappedToDo>();
		default:
			// Any other case is either handled by an add-in or not at all
			return std::make_shared<addinParser>(type);
		}
	}

	_Check_return_ ULONG BuildFlagIndexFromTag(
		ULONG ulPropTag,
		ULONG ulPropNameID,
		_In_opt_z_ LPCWSTR lpszPropNameString,
		_In_opt_ LPCGUID lpguidNamedProp) noexcept
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

	_Check_return_ parserType FindSmartViewParserForProp(
		const ULONG ulPropTag,
		const ULONG ulPropNameID,
		_In_opt_ const LPCGUID lpguidNamedProp) noexcept
	{
		const auto ulIndex = BuildFlagIndexFromTag(ulPropTag, ulPropNameID, nullptr, lpguidNamedProp);
		const auto bMV = (PROP_TYPE(ulPropTag) & MV_FLAG) == MV_FLAG;

		for (const auto& block : SmartViewParserArray)
		{
			if (block.ulIndex == ulIndex && block.bMV == bMV) return block.type;
		}

		return parserType::NOPARSING;
	}

	_Check_return_ parserType FindSmartViewParserForProp(
		const ULONG ulPropTag,
		const ULONG ulPropNameID,
		_In_opt_ const LPCGUID lpguidNamedProp,
		bool bMVRow) noexcept
	{
		auto ulLookupPropTag = ulPropTag;
		if (bMVRow) ulLookupPropTag |= MV_FLAG;

		return FindSmartViewParserForProp(ulLookupPropTag, ulPropNameID, lpguidNamedProp);
	}

	_Check_return_ parserType FindSmartViewParserForProp(
		_In_opt_ const _SPropValue* lpProp, // required property value
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID* lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const SBinary* lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
		bool bMVRow) // did the row come from a MV prop?
	{
		if (!lpProp) return parserType::NOPARSING;
		return FindSmartViewParserForProp(lpProp->ulPropTag, lpMAPIProp, lpNameID, lpMappingSignature, bIsAB, bMVRow);
	}

	_Check_return_ parserType FindSmartViewParserForProp(
		_In_opt_ const ULONG ulPropTag, // required property value
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID* lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const SBinary* lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
		bool bMVRow) // did the row come from a MV prop?
	{
		if (!registry::doSmartView) return parserType::NOPARSING;

		const auto npi = GetNamedPropInfo(ulPropTag, lpMAPIProp, lpNameID, lpMappingSignature, bIsAB);
		const auto ulPropNameID = npi.first;
		const auto propNameGUID = npi.second;

		switch (PROP_TYPE(ulPropTag))
		{
		case PT_BINARY:
		{
			auto ulLookupPropTag = ulPropTag;
			if (bMVRow) ulLookupPropTag |= MV_FLAG;

			auto parser = FindSmartViewParserForProp(ulLookupPropTag, ulPropNameID, &propNameGUID);
			// We special-case this property
			if (parser != parserType::NOPARSING && PR_ROAMING_BINARYSTREAM == ulLookupPropTag && lpMAPIProp)
			{
				auto lpPropSubject = LPSPropValue{};

				WC_MAPI_S(HrGetOneProp(lpMAPIProp, PR_SUBJECT_W, &lpPropSubject));

				if (strings::CheckStringProp(lpPropSubject, PT_UNICODE) &&
					0 == wcscmp(lpPropSubject->Value.lpszW, L"IPM.Configuration.Autocomplete"))
				{
					parser = parserType::NICKNAMECACHE;
				}

				MAPIFreeBuffer(lpPropSubject);
			}

			return parser;
		}

		case PT_MV_BINARY:
			return FindSmartViewParserForProp(ulPropTag, ulPropNameID, &propNameGUID);
		}

		return parserType::NOPARSING;
	}

	_Check_return_ std::pair<ULONG, GUID> GetNamedPropInfo(_In_opt_ const MAPINAMEID* lpNameID) noexcept
	{
		auto ulPropNameID = ULONG{};
		auto propNameGUID = GUID{};
		if (lpNameID)
		{
			if (lpNameID->lpguid)
			{
				propNameGUID = *lpNameID->lpguid;
			}

			if (lpNameID->ulKind == MNID_ID)
			{
				ulPropNameID = lpNameID->Kind.lID;
			}
		}

		return {ulPropNameID, propNameGUID};
	}

	std::pair<ULONG, GUID> GetNamedPropInfo(
		_In_opt_ ULONG ulPropTag,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const MAPINAMEID* lpNameID,
		_In_opt_ const SBinary* lpMappingSignature,
		bool bIsAB)
	{
		if (lpNameID) return GetNamedPropInfo(lpNameID);

		if (lpMAPIProp && // if we have an object
			!bIsAB && // Some address book providers return garbage which will crash us
			registry::parseNamedProps && // and we're parsing named props
			(registry::getPropNamesOnAllProps ||
			 PROP_ID(ulPropTag) >= 0x8000)) // and it's either a named prop or we're doing all props
		{
			const auto name = cache::GetNameFromID(lpMAPIProp, lpMappingSignature, ulPropTag, 0);
			if (cache::namedPropCacheEntry::valid(name))
			{
				return GetNamedPropInfo(name->getMapiNameId());
			}
		}

		return {};
	}

	std::wstring parsePropertySmartView(
		_In_opt_ const SPropValue* lpProp, // required property value
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID* lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const SBinary* lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
		bool bMVRow) // did the row come from a MV prop?
	{
		if (!registry::doSmartView || !lpProp) return strings::emptystring;

		const auto npi = GetNamedPropInfo(lpProp->ulPropTag, lpMAPIProp, lpNameID, lpMappingSignature, bIsAB);
		const auto ulPropNameID = npi.first;
		const auto propNameGUID = npi.second;

		const auto parser = FindSmartViewParserForProp(lpProp, lpMAPIProp, lpNameID, lpMappingSignature, bIsAB, bMVRow);

		ULONG ulLookupPropTag = NULL;
		switch (PROP_TYPE(lpProp->ulPropTag))
		{
		case PT_LONG:
			return InterpretNumberAsString(
				lpProp->Value.l, lpProp->ulPropTag, ulPropNameID, nullptr, &propNameGUID, true);
		case PT_I2:
			return InterpretNumberAsString(
				lpProp->Value.i, lpProp->ulPropTag, ulPropNameID, nullptr, &propNameGUID, true);
		case PT_I8:
			return InterpretNumberAsString(
				lpProp->Value.li.QuadPart, lpProp->ulPropTag, ulPropNameID, nullptr, &propNameGUID, true);
		case PT_MV_LONG:
			return InterpretMVLongAsString(lpProp->Value.MVl, lpProp->ulPropTag, ulPropNameID, &propNameGUID);
		case PT_BINARY:
			// TODO: Find a way to use ulLookupPropTag here
			ulLookupPropTag = lpProp->ulPropTag;
			if (bMVRow) ulLookupPropTag |= MV_FLAG;

			if (parser != parserType::NOPARSING)
			{
				return InterpretBinary(mapi::getBin(lpProp), parser, lpMAPIProp)->toString();
			}

			break;
		case PT_MV_BINARY:
			if (parser != parserType::NOPARSING)
			{
				return InterpretMVBinaryAsString(lpProp->Value.MVbin, parser, lpMAPIProp);
			}

			break;
		}

		return strings::emptystring;
	}

	std::wstring InterpretMVBinaryAsString(SBinaryArray myBinArray, parserType parser, _In_opt_ LPMAPIPROP lpMAPIProp)
	{
		if (!registry::doSmartView) return L"";

		std::wstring szResult;

		for (ULONG ulRow = 0; ulRow < myBinArray.cValues; ulRow++)
		{
			if (ulRow != 0)
			{
				szResult += L"\r\n"; // STRING_OK
			}

			szResult += strings::formatmessage(IDS_MVROWBIN, ulRow);
			szResult += InterpretBinary(myBinArray.lpbin[ulRow], parser, lpMAPIProp)->toString();
		}

		return szResult;
	}

	std::wstring InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag)
	{
		return InterpretNumberAsString(ulVal, ulPropTag, NULL, nullptr, nullptr, false);
	}

	std::wstring InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp)
	{
		return InterpretNumberAsString(ulVal, PT_LONG, ulPropNameID, nullptr, lpguidNamedProp, false);
	}

	// Interprets a LONGLONG and returns a string
	// Will not return a string if the lpProp is not a PT_LONG/PT_I2/PT_I8 or we don't recognize the property
	// Will use named property details to look up named property flags
	std::wstring InterpretNumberAsString(
		LONGLONG val,
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
		case parserType::LONGRTIME:
			lpszResultString = RTimeToSzString(static_cast<DWORD>(val), bLabel);
			break;
		case parserType::PTI8:
		{
			auto li = LARGE_INTEGER{};
			li.QuadPart = val;
			lpszResultString = PTI8ToSzString(li, bLabel);
		}
		break;
		case parserType::SFIDMID:
			lpszResultString = FidMidToSzString(val, bLabel);
			break;
			// insert future parsers here
		default:
			ulPropID = BuildFlagIndexFromTag(ulPropTag, ulPropNameID, lpszPropNameString, lpguidNamedProp);
			if (ulPropID)
			{
				lpszResultString += flags::InterpretFlags(ulPropID, static_cast<LONG>(val));

				if (bLabel && !lpszResultString.empty())
				{
					lpszResultString = strings::formatmessage(IDS_FLAGS_PREFIX) + lpszResultString;
				}
			}

			break;
		}

		return lpszResultString;
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
		switch (PROP_TYPE(ulPropTag))
		{
		case PT_LONG:
			return InterpretNumberAsString(pV.ul, ulPropTag, ulPropNameID, lpszPropNameString, lpguidNamedProp, bLabel);
		case PT_I2:
			return InterpretNumberAsString(pV.i, ulPropTag, ulPropNameID, lpszPropNameString, lpguidNamedProp, bLabel);
		case PT_I8:
			return InterpretNumberAsString(
				pV.li.QuadPart, ulPropTag, ulPropNameID, lpszPropNameString, lpguidNamedProp, bLabel);
		}

		return strings::emptystring;
	}

	std::wstring InterpretMVLongAsString(
		std::vector<LONG> rows,
		_In_opt_ ULONG ulPropTag,
		_In_opt_ ULONG ulPropNameID,
		_In_opt_ LPGUID lpguidNamedProp)
	{
		if (!registry::doSmartView) return strings::emptystring;

		auto szArray = std::vector<std::wstring>{};
		ULONG ulRow = 0;

		for (const auto& row : rows)
		{
			const auto szSmartView = InterpretNumberAsString(
				row, CHANGE_PROP_TYPE(ulPropTag, PT_LONG), ulPropNameID, nullptr, lpguidNamedProp, true);
			if (!szSmartView.empty())
			{
				szArray.push_back(strings::formatmessage(IDS_MVROWLONG, ulRow++, szSmartView.c_str()));
			}
		}

		return strings::join(szArray, L"\r\n");
	}

	std::wstring InterpretMVLongAsString(
		SLongArray myLongArray,
		_In_opt_ ULONG ulPropTag,
		_In_opt_ ULONG ulPropNameID,
		_In_opt_ LPCGUID lpguidNamedProp)
	{
		if (!registry::doSmartView) return strings::emptystring;

		auto szArray = std::vector<std::wstring>{};

		for (ULONG ulRow = 0; ulRow < myLongArray.cValues; ulRow++)
		{
			auto szSmartView = InterpretNumberAsString(
				myLongArray.lpl[ulRow],
				CHANGE_PROP_TYPE(ulPropTag, PT_LONG),
				ulPropNameID,
				nullptr,
				lpguidNamedProp,
				true);
			if (!szSmartView.empty())
			{
				szArray.push_back(strings::formatmessage(IDS_MVROWLONG, ulRow, szSmartView.c_str()));
			}
		}

		return strings::join(szArray, L"\r\n");
	}

	_Check_return_ std::wstring RTimeToString(DWORD rTime)
	{
		std::wstring rTimeString;
		std::wstring rTimeAltString;
		auto fTime = FILETIME{};
		auto liNumSec = LARGE_INTEGER{};
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

	_Check_return_ constexpr WORD WGetReplId(ID id) noexcept { return id.replid; }

	_Check_return_ constexpr ULONGLONG UllGetIdGlobcnt(ID id) noexcept
	{
		ULONGLONG ul = 0;
		for (const auto i : id.globcnt)
		{
			ul <<= 8;
			ul += i;
		}

		return ul;
	}

	_Check_return_ std::wstring FidMidToSzString(LONGLONG llID, bool bLabel)
	{
		const auto pid = reinterpret_cast<const ID*>(&llID);
		if (bLabel)
		{
			return strings::formatmessage(IDS_FIDMIDFORMATLABEL, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
		}

		return strings::formatmessage(IDS_FIDMIDFORMAT, WGetReplId(*pid), UllGetIdGlobcnt(*pid));
	}
} // namespace smartview