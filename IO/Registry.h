// Header file for the MFCMAPI application registry functions

#pragma once

namespace registry
{
#define RKEY_ROOT L"SOFTWARE\\Microsoft\\MFCMAPI" // STRING_OK

	enum __REGKEYTYPES
	{
		regDWORD,
		regSTRING
	};

	enum __REGOPTIONTYPE
	{
		regoptCheck,
		regoptString,
		regoptStringHex,
		regoptStringDec
	};

	class __RegKey
	{
	public:
		std::wstring szKeyName;
		__REGKEYTYPES ulRegKeyType;
		__REGOPTIONTYPE ulRegOptType;
		DWORD ulDefDWORD;
		DWORD ulCurDWORD;
		std::wstring szDefSTRING;
		std::wstring szCurSTRING;
		bool bRefresh;
		UINT uiOptionsPrompt;
	};

	class boolRegKey : public __RegKey
	{
	public:
		boolRegKey(const std::wstring& _szKeyName, bool _default, bool _refresh, int _uiOptionsPrompt)
		{
			szKeyName = _szKeyName;
			ulRegKeyType = regDWORD;
			ulRegOptType = regoptCheck;
			ulDefDWORD = _default;
			bRefresh = _refresh;
			uiOptionsPrompt = _uiOptionsPrompt;
		}

		operator bool() const { return ulCurDWORD != 0; }

		boolRegKey& operator=(DWORD val)
		{
			ulCurDWORD = val;
			return *this;
		}
	};

	class dwordRegKey : public __RegKey
	{
	public:
		dwordRegKey(
			const std::wstring& _szKeyName,
			__REGOPTIONTYPE _ulRegOptType,
			DWORD _default,
			bool _refresh,
			int _uiOptionsPrompt)
		{
			szKeyName = _szKeyName;
			ulRegKeyType = regDWORD;
			ulRegOptType = _ulRegOptType;
			ulDefDWORD = _default;
			bRefresh = _refresh;
			uiOptionsPrompt = _uiOptionsPrompt;
		}

		operator DWORD() const { return ulCurDWORD; }

		dwordRegKey& operator=(DWORD val)
		{
			ulCurDWORD = val;
			return *this;
		}

		dwordRegKey& operator|=(DWORD val)
		{
			ulCurDWORD |= val;
			return *this;
		}
	};

	class wstringRegKey : public __RegKey
	{
	public:
		wstringRegKey(const std::wstring& _szKeyName, const std::wstring& _default, bool _refresh, int _uiOptionsPrompt)
		{
			szKeyName = _szKeyName;
			ulRegKeyType = regSTRING;
			ulRegOptType = regoptString;
			szCurSTRING = _default;
			bRefresh = _refresh;
			uiOptionsPrompt = _uiOptionsPrompt;
		}

		operator std::wstring() const { return szCurSTRING; }

		wstringRegKey& operator=(std::wstring val)
		{
			szCurSTRING = val;
			return *this;
		}

		_NODISCARD std::wstring::size_type length() const noexcept { return szCurSTRING.length(); }
		_NODISCARD bool empty() const noexcept { return szCurSTRING.empty(); }
		void clear() noexcept { szCurSTRING.clear(); }
		void push_back(const wchar_t _Ch) { szCurSTRING.push_back(_Ch); }
		_NODISCARD std::wstring::reference operator[](const std::wstring::size_type _Off) { return szCurSTRING[_Off]; }
	};

	// Registry key Names
	enum REGKEYNAMES
	{
		regkeyDEBUG_TAG,
		regkeyDEBUG_TO_FILE,
		regkeyDEBUG_FILE_NAME,
		regkeyPARSED_NAMED_PROPS,
		regkeyGETPROPNAMES_ON_ALL_PROPS,
		regkeyTHROTTLE_LEVEL,
		regkeyHIER_EXPAND_NOTIFS,
		regkeyHIER_ROOT_NOTIFS,
		regkeyDO_SMART_VIEW,
		regkeyONLY_ADDITIONAL_PROPERTIES,
		regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST,
		regkeyUSE_GETPROPLIST,
		regkeyPREFER_UNICODE_PROPS,
		regkeyCACHE_NAME_DPROPS,
		regkeyALLOW_DUPE_COLUMNS,
		regkeyDO_COLUMN_NAMES,
		regkeyEDIT_COLUMNS_ON_LOAD,
		regkeyMDB_ONLINE,
		regKeyMAPI_NO_CACHE,
		regkeyALLOW_PERSIST_CACHE,
		regkeyUSE_IMAPIPROGRESS,
		regkeyUSE_MESSAGERAW,
		regkeySUPPRESS_NOT_FOUND,
		regkeyHEAPENABLETERMINATIONONCORRUPTION,
		regkeyLOADADDINS,
		regkeyFORCEOUTLOOKMAPI,
		regkeyFORCESYSTEMMAPI,
		regkeyHEX_DIALOG_DIAG,
		regkeyDISPLAY_ABOUT_DIALOG,
		regkeyPROP_COLUMN_ORDER,
		NUMRegKeys
	};

#define NumRegOptionKeys (registry::NUMRegKeys - 2)

	extern __RegKey RegKeys[NUMRegKeys];

	void SetDefaults();
	void WriteToRegistry();
	void ReadFromRegistry();

	_Check_return_ HKEY CreateRootKey();

	DWORD ReadDWORDFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ DWORD dwDefaultVal = 0);
	std::wstring ReadStringFromRegistry(
		_In_ HKEY hKey,
		_In_ const std::wstring& szValue,
		_In_ const std::wstring& szDefault = strings::emptystring);

	void WriteStringToRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValueName, _In_ const std::wstring& szValue);

	// Registry settings
#ifdef _DEBUG
	static auto debugTag = dwordRegKey{L"DebugTag", regoptStringHex, DBGAll, false, IDS_REGKEY_DEBUG_TAG};
#else
	static auto debugTag = dwordRegKey{L"DebugTag", regoptStringHex, DBGNoDebug, 0, false, IDS_REGKEY_DEBUG_TAG};
#endif
	static auto debugToFile = boolRegKey{L"DebugToFile", false, false, IDS_REGKEY_DEBUG_TO_FILE};
	static auto debugFileName = wstringRegKey{L"DebugFileName", L"c:\\mfcmapi.log", false, IDS_REGKEY_DEBUG_FILE_NAME};
	static auto getPropNamesOnAllProps =
		boolRegKey{L"GetPropNamesOnAllProps", false, true, IDS_REGKEY_GETPROPNAMES_ON_ALL_PROPS};
	static auto parseNamedProps = boolRegKey{L"ParseNamedProps", false, true, IDS_REGKEY_PARSED_NAMED_PROPS};
	static auto throttleLevel = dwordRegKey{L"ThrottleLevel", regoptStringDec, 0, false, IDS_REGKEY_THROTTLE_LEVEL};
	static auto hierExpandNotifications =
		boolRegKey{L"HierExpandNotifications", true, false, IDS_REGKEY_HIER_EXPAND_NOTIFS};
	static auto hierRootNotifs = boolRegKey{L"HierRootNotifs", false, false, IDS_REGKEY_HIER_ROOT_NOTIFS};
	static auto doSmartView = boolRegKey{L"DoSmartView", true, true, IDS_REGKEY_DO_SMART_VIEW};
	static auto onlyAdditionalProperties =
		boolRegKey{L"OnlyAdditionalProperties", false, true, IDS_REGKEY_ONLYADDITIONALPROPERTIES};
	static auto useRowDataForSinglePropList =
		boolRegKey{L"UseRowDataForSinglePropList", false, true, IDS_REGKEY_USE_ROW_DATA_FOR_SINGLEPROPLIST};
	static auto useGetPropList = boolRegKey{L"UseGetPropList", true, true, IDS_REGKEY_USE_GETPROPLIST};
	static auto preferUnicodeProps = boolRegKey{L"PreferUnicodeProps", true, true, IDS_REGKEY_PREFER_UNICODE_PROPS};
	static auto cacheNamedProps = boolRegKey{L"CacheNamedProps", true, false, IDS_REGKEY_CACHE_NAMED_PROPS};
	static auto allowDupeColumns = boolRegKey{L"AllowDupeColumns", false, false, IDS_REGKEY_ALLOW_DUPE_COLUMNS};
	static auto doColumnNames = boolRegKey{L"DoColumnNames", true, false, IDS_REGKEY_DO_COLUMN_NAMES};
	static auto editColumnsOnLoad = boolRegKey{L"EditColumnsOnLoad", false, false, IDS_REGKEY_EDIT_COLUMNS_ON_LOAD};
	static auto forceMDBOnline = boolRegKey{L"ForceMDBOnline", false, false, IDS_REGKEY_MDB_ONLINE};
	static auto forceMapiNoCache = boolRegKey{L"ForceMapiNoCache", false, false, IDS_REGKEY_MAPI_NO_CACHE};
	static auto allowPersistCache = boolRegKey{L"AllowPersistCache", false, false, IDS_REGKEY_ALLOW_PERSIST_CACHE};
	static auto useIMAPIProgress = boolRegKey{L"UseIMAPIProgress", false, false, IDS_REGKEY_USE_IMAPIPROGRESS};
	static auto useMessageRaw = boolRegKey{L"UseMessageRaw", false, false, IDS_REGKEY_USE_MESSAGERAW};
	static auto suppressNotFound = boolRegKey{L"SuppressNotFound", true, false, IDS_REGKEY_SUPPRESS_NOTFOUND};
	static auto heapEnableTerminationOnCorruption =
		boolRegKey{L"HeapEnableTerminationOnCorruption", true, false, IDS_REGKEY_HEAPENABLETERMINATIONONCORRUPTION};
	static auto loadAddIns = boolRegKey{L"LoadAddIns", true, false, IDS_REGKEY_LOADADDINS};
	static auto forceOutlookMAPI = boolRegKey{L"ForceOutlookMAPI", false, false, IDS_REGKEY_FORCEOUTLOOKMAPI};
	static auto forceSystemMAPI = boolRegKey{L"ForceSystemMAPI", false, false, IDS_REGKEY_FORCESYSTEMMAPI};
	static auto hexDialogDiag = boolRegKey{L"HexDialogDiag", false, false, IDS_REGKEY_HEXDIALOGDIAG};
	static auto displayAboutDialog = boolRegKey{L"DisplayAboutDialog", true, false, NULL};
	static auto propertyColumnOrder = wstringRegKey{L"PropertyColumnOrder", L"", false, NULL};
} // namespace registry