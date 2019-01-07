#include <StdAfx.h>
#include <IO/Registry.h>

namespace registry
{
	// Keep this in sync with REGKEYNAMES
	// clang-format off
	__RegKeys RegKeys[] = {
	#ifdef _DEBUG
		{ L"DebugTag",					regDWORD,regoptStringHex,	DBGAll		,0,L"",L"",false,	IDS_REGKEY_DEBUG_TAG }, // STRING_OK
	#else
		{ L"DebugTag",					regDWORD,regoptStringHex,	DBGNoDebug	,0,L"",L"",false,	IDS_REGKEY_DEBUG_TAG }, // STRING_OK
	#endif
		debugToFile,
		{ L"DebugFileName",				regSTRING,regoptString,		0			,0,L"c:\\mfcmapi.log",L"",false,	IDS_REGKEY_DEBUG_FILE_NAME }, // STRING_OK
		{ L"ParseNamedProps",			regDWORD,regoptCheck,		true		,0,L"",L"",true,	IDS_REGKEY_PARSED_NAMED_PROPS }, // STRING_OK
		{ L"GetPropNamesOnAllProps",	regDWORD,regoptCheck,		false		,0,L"",L"",true,	IDS_REGKEY_GETPROPNAMES_ON_ALL_PROPS }, // STRING_OK
		{ L"ThrottleLevel",				regDWORD,regoptStringDec,	0			,0,L"",L"",false,	IDS_REGKEY_THROTTLE_LEVEL }, // STRING_OK
		{ L"HierExpandNotifications",	regDWORD,regoptCheck,		true		,0,L"",L"",false,	IDS_REGKEY_HIER_EXPAND_NOTIFS }, // STRING_OK
		{ L"HierRootNotifs",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_HIER_ROOT_NOTIFS }, // STRING_OK
		{ L"DoSmartView",				regDWORD,regoptCheck,		true		,0,L"",L"",true,	IDS_REGKEY_DO_SMART_VIEW }, // STRING_OK
		{ L"OnlyAdditionalProperties",	regDWORD,regoptCheck,		false		,0,L"",L"",true,	IDS_REGKEY_ONLYADDITIONALPROPERTIES }, // STRING_OK
		{ L"UseRowDataForSinglePropList",	regDWORD,regoptCheck,	false		,0,L"",L"",true,	IDS_REGKEY_USE_ROW_DATA_FOR_SINGLEPROPLIST }, // STRING_OK
		{ L"UseGetPropList",			regDWORD,regoptCheck,		true		,0,L"",L"",true,	IDS_REGKEY_USE_GETPROPLIST }, // STRING_OK
		{ L"PreferUnicodeProps",		regDWORD,regoptCheck,		true		,0,L"",L"",true,	IDS_REGKEY_PREFER_UNICODE_PROPS }, // STRING_OK
		{ L"CacheNamedProps",			regDWORD,regoptCheck,		true		,0,L"",L"",false,	IDS_REGKEY_CACHE_NAMED_PROPS }, // STRING_OK
		{ L"AllowDupeColumns",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_ALLOW_DUPE_COLUMNS }, // STRING_OK
		{ L"DoColumnNames",				regDWORD,regoptCheck,		true		,0,L"",L"",false,	IDS_REGKEY_DO_COLUMN_NAMES }, // STRING_OK
		{ L"EditColumnsOnLoad",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_EDIT_COLUMNS_ON_LOAD }, // STRING_OK
		{ L"ForceMDBOnline",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_MDB_ONLINE }, // STRING_OK
		{ L"ForceMapiNoCache",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_MAPI_NO_CACHE }, // STRING_OK
		{ L"AllowPersistCache",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_ALLOW_PERSIST_CACHE }, // STRING_OK
		{ L"UseIMAPIProgress",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_USE_IMAPIPROGRESS }, // STRING_OK
		{ L"UseMessageRaw",				regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_USE_MESSAGERAW }, // STRING_OK
		{ L"SuppressNotFound",			regDWORD,regoptCheck,		true		,0,L"",L"",false,	IDS_REGKEY_SUPPRESS_NOTFOUND }, // STRING_OK
		{ L"HeapEnableTerminationOnCorruption",regDWORD,regoptCheck,true		,0,L"",L"",false,	IDS_REGKEY_HEAPENABLETERMINATIONONCORRUPTION }, // STRING_OK
		{ L"LoadAddIns",				regDWORD,regoptCheck,		true		,0,L"",L"",false,	IDS_REGKEY_LOADADDINS }, // STRING_OK
		{ L"ForceOutlookMAPI",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_FORCEOUTLOOKMAPI }, // STRING_OK
		{ L"ForceSystemMAPI",			regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_FORCESYSTEMMAPI }, // STRING_OK
		{ L"HexDialogDiag",				regDWORD,regoptCheck,		false		,0,L"",L"",false,	IDS_REGKEY_HEXDIALOGDIAG }, // STRING_OK
		{ L"DisplayAboutDialog",		regDWORD,regoptCheck,		true		,0,L"",L"",false,	NULL }, // STRING_OK
		{ L"PropertyColumnOrder",		regSTRING,regoptCheck,		0			,0,L"",L"",false,	NULL }, // STRING_OK
		// {KeyName,					keytype,opttype,			defaultDWORD,0,defaultString,NULL,bRefresh,IDS_REGKEY_*} // Regkey template
	};
	// clang-format on

	void SetDefaults()
	{
		// Set some defaults to begin with:
		for (auto& regKey : RegKeys)
		{
			if (regKey.ulRegKeyType == regDWORD)
			{
				regKey.ulCurDWORD = regKey.ulDefDWORD;
			}
			else if (regKey.ulRegKeyType == regSTRING)
			{
				regKey.szCurSTRING = regKey.szDefSTRING;
			}
		}
	}

	// $--HrGetRegistryValue---------------------------------------------------------
	// Get a registry value - allocating memory using new to hold it.
	// -----------------------------------------------------------------------------
	_Check_return_ HRESULT HrGetRegistryValue(
		_In_ HKEY hKey, // the key.
		_In_ const std::wstring& lpszValue, // value name in key.
		_Out_ DWORD* lpType, // where to put type info.
		_Out_ LPVOID* lppData) // where to put the data.
	{
		output::DebugPrint(DBGGeneric, L"HrGetRegistryValue(%ws)\n", lpszValue.c_str());

		*lppData = nullptr;
		DWORD cb = NULL;

		// Get its size
		auto hRes = WC_W32(RegQueryValueExW(hKey, lpszValue.c_str(), nullptr, lpType, nullptr, &cb));

		// only handle types we know about - all others are bad
		if (hRes == S_OK && cb && (REG_SZ == *lpType || REG_DWORD == *lpType || REG_MULTI_SZ == *lpType))
		{
			*lppData = new BYTE[cb];

			if (*lppData)
			{
				// Get the current value
				hRes = EC_W32(RegQueryValueExW(
					hKey, lpszValue.c_str(), nullptr, lpType, static_cast<unsigned char*>(*lppData), &cb));

				if (FAILED(hRes))
				{
					delete[] * lppData;
					*lppData = nullptr;
				}
			}
		}
		else
			hRes = MAPI_E_INVALID_PARAMETER;

		return hRes;
	}

	// If the value is not set in the registry, return the default value
	DWORD ReadDWORDFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ DWORD dwDefaultVal)
	{
		if (szValue.empty()) return dwDefaultVal;
		DWORD dwKeyType = NULL;
		DWORD* lpValue = nullptr;
		auto ret = dwDefaultVal;

		const auto hRes = WC_H(HrGetRegistryValue(hKey, szValue, &dwKeyType, reinterpret_cast<LPVOID*>(&lpValue)));
		if (hRes == S_OK && REG_DWORD == dwKeyType && lpValue)
		{
			ret = *lpValue;
		}

		delete[] lpValue;
		return ret;
	}
	// If the value is not set in the registry, return the default value
	std::wstring
	ReadStringFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ const std::wstring& szDefault)
	{
		if (szValue.empty()) return szDefault;
		output::DebugPrint(DBGGeneric, L"ReadStringFromRegistry(%ws)\n", szValue.c_str());

		DWORD dwKeyType = NULL;
		LPBYTE szBuf = nullptr;
		DWORD cb = NULL;

		// Get its size
		auto hRes = WC_W32(RegQueryValueExW(hKey, szValue.c_str(), nullptr, &dwKeyType, nullptr, &cb));

		if (hRes == S_OK && cb && !(cb % 2) && REG_SZ == dwKeyType)
		{
			szBuf = new (std::nothrow) BYTE[cb];
			if (szBuf)
			{
				// Get the current value
				hRes = EC_W32(RegQueryValueExW(hKey, szValue.c_str(), nullptr, &dwKeyType, szBuf, &cb));

				if (FAILED(hRes))
				{
					delete[] szBuf;
					szBuf = nullptr;
				}
			}
		}

		auto ret = szDefault;
		if (hRes == S_OK && cb && !(cb % 2) && REG_SZ == dwKeyType && szBuf)
		{
			ret = std::wstring(LPWSTR(szBuf), cb / sizeof WCHAR);
		}

		delete[] szBuf;
		return ret;
	}

	void ReadFromRegistry()
	{
		HKEY hRootKey = nullptr;
		WC_W32_S(RegOpenKeyExW(HKEY_CURRENT_USER, RKEY_ROOT, NULL, KEY_READ, &hRootKey));

		// Now that we have a root key, go get our values
		if (hRootKey)
		{
			for (auto& regKey : RegKeys)
			{
				if (regKey.ulRegKeyType == regDWORD)
				{
					regKey.ulCurDWORD = ReadDWORDFromRegistry(hRootKey, regKey.szKeyName, regKey.ulCurDWORD);
				}
				else if (regKey.ulRegKeyType == regSTRING)
				{
					regKey.szCurSTRING = ReadStringFromRegistry(hRootKey, regKey.szKeyName, regKey.szCurSTRING);
				}
			}

			EC_W32_S(RegCloseKey(hRootKey));
		}

		output::SetDebugLevel(debugTag);
		output::DebugPrintVersion(DBGVersionBanner);
	}

	void WriteDWORDToRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValueName, DWORD dwValue)
	{
		WC_W32_S(RegSetValueExW(
			hKey, szValueName.c_str(), NULL, REG_DWORD, reinterpret_cast<LPBYTE>(&dwValue), sizeof(DWORD)));
	}

	void CommitDWORDIfNeeded(_In_ HKEY hKey, _In_ const std::wstring& szValueName, DWORD dwValue, DWORD dwDefaultValue)
	{
		if (dwValue != dwDefaultValue)
		{
			WriteDWORDToRegistry(hKey, szValueName, dwValue);
		}
		else
		{
			WC_W32_S(RegDeleteValueW(hKey, szValueName.c_str()));
		}
	}

	void WriteStringToRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValueName, _In_ const std::wstring& szValue)
	{
		// Reg needs bytes, so CB is correct here
		auto cbValue = szValue.length() * sizeof(WCHAR);
		cbValue += sizeof(WCHAR); // NULL terminator

		WC_W32_S(RegSetValueExW(
			hKey, szValueName.c_str(), NULL, REG_SZ, LPBYTE(szValue.c_str()), static_cast<DWORD>(cbValue)));
	}

	void CommitStringIfNeeded(
		_In_ HKEY hKey,
		_In_ const std::wstring& szValueName,
		_In_ const std::wstring& szValue,
		_In_ const std::wstring& szDefaultValue)
	{
		if (strings::wstringToLower(szValue) != strings::wstringToLower(szDefaultValue))
		{
			WriteStringToRegistry(hKey, szValueName, szValue);
		}
		else
		{
			WC_W32_S(RegDeleteValueW(hKey, szValueName.c_str()));
		}
	}

	_Check_return_ HKEY CreateRootKey()
	{
		HKEY hkSub = nullptr;

		// Try to open the root key before we do the work to create it
		const auto hRes = WC_W32(RegOpenKeyExW(HKEY_CURRENT_USER, RKEY_ROOT, NULL, KEY_READ | KEY_WRITE, &hkSub));
		if (SUCCEEDED(hRes) && hkSub) return hkSub;

		WC_W32_S(RegCreateKeyExW(
			HKEY_CURRENT_USER, RKEY_ROOT, 0, nullptr, 0, KEY_READ | KEY_WRITE, nullptr, &hkSub, nullptr));

		return hkSub;
	}

	void WriteToRegistry()
	{
		const auto hRootKey = CreateRootKey();

		// Now that we have a root key, go set our values
		for (auto& regKey : RegKeys)
		{
			if (regKey.ulRegKeyType == regDWORD)
			{
				CommitDWORDIfNeeded(hRootKey, regKey.szKeyName, regKey.ulCurDWORD, regKey.ulDefDWORD);
			}
			else if (regKey.ulRegKeyType == regSTRING)
			{
				CommitStringIfNeeded(hRootKey, regKey.szKeyName, regKey.szCurSTRING, regKey.szDefSTRING);
			}
		}

		if (hRootKey) EC_W32_S(RegCloseKey(hRootKey));
	}
} // namespace registry