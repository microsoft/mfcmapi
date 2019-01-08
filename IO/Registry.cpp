#include <StdAfx.h>
#include <IO/Registry.h>

namespace registry
{
	// Registry settings
#ifdef _DEBUG
	dwordRegKey debugTag{L"DebugTag", regoptStringHex, DBGAll, false, IDS_REGKEY_DEBUG_TAG};
#else
	dwordRegKey debugTag{L"DebugTag", regoptStringHex, DBGNoDebug, false, IDS_REGKEY_DEBUG_TAG};
#endif
	boolRegKey debugToFile{L"DebugToFile", false, false, IDS_REGKEY_DEBUG_TO_FILE};
	wstringRegKey debugFileName{L"DebugFileName", L"c:\\mfcmapi.log", false, IDS_REGKEY_DEBUG_FILE_NAME};
	boolRegKey getPropNamesOnAllProps{L"GetPropNamesOnAllProps", false, true, IDS_REGKEY_GETPROPNAMES_ON_ALL_PROPS};
	boolRegKey parseNamedProps{L"ParseNamedProps", false, true, IDS_REGKEY_PARSED_NAMED_PROPS};
	dwordRegKey throttleLevel{L"ThrottleLevel", regoptStringDec, 0, false, IDS_REGKEY_THROTTLE_LEVEL};
	boolRegKey hierExpandNotifications{L"HierExpandNotifications", true, false, IDS_REGKEY_HIER_EXPAND_NOTIFS};
	boolRegKey hierRootNotifs{L"HierRootNotifs", false, false, IDS_REGKEY_HIER_ROOT_NOTIFS};
	boolRegKey doSmartView{L"DoSmartView", true, true, IDS_REGKEY_DO_SMART_VIEW};
	boolRegKey onlyAdditionalProperties{L"OnlyAdditionalProperties", false, true, IDS_REGKEY_ONLYADDITIONALPROPERTIES};
	boolRegKey useRowDataForSinglePropList{L"UseRowDataForSinglePropList",
										   false,
										   true,
										   IDS_REGKEY_USE_ROW_DATA_FOR_SINGLEPROPLIST};
	boolRegKey useGetPropList{L"UseGetPropList", true, true, IDS_REGKEY_USE_GETPROPLIST};
	boolRegKey preferUnicodeProps{L"PreferUnicodeProps", true, true, IDS_REGKEY_PREFER_UNICODE_PROPS};
	boolRegKey cacheNamedProps{L"CacheNamedProps", true, false, IDS_REGKEY_CACHE_NAMED_PROPS};
	boolRegKey allowDupeColumns{L"AllowDupeColumns", false, false, IDS_REGKEY_ALLOW_DUPE_COLUMNS};
	boolRegKey doColumnNames{L"DoColumnNames", true, false, IDS_REGKEY_DO_COLUMN_NAMES};
	boolRegKey editColumnsOnLoad{L"EditColumnsOnLoad", false, false, IDS_REGKEY_EDIT_COLUMNS_ON_LOAD};
	boolRegKey forceMDBOnline{L"ForceMDBOnline", false, false, IDS_REGKEY_MDB_ONLINE};
	boolRegKey forceMapiNoCache{L"ForceMapiNoCache", false, false, IDS_REGKEY_MAPI_NO_CACHE};
	boolRegKey allowPersistCache{L"AllowPersistCache", false, false, IDS_REGKEY_ALLOW_PERSIST_CACHE};
	boolRegKey useIMAPIProgress{L"UseIMAPIProgress", false, false, IDS_REGKEY_USE_IMAPIPROGRESS};
	boolRegKey useMessageRaw{L"UseMessageRaw", false, false, IDS_REGKEY_USE_MESSAGERAW};
	boolRegKey suppressNotFound{L"SuppressNotFound", true, false, IDS_REGKEY_SUPPRESS_NOTFOUND};
	boolRegKey heapEnableTerminationOnCorruption{L"HeapEnableTerminationOnCorruption",
												 true,
												 false,
												 IDS_REGKEY_HEAPENABLETERMINATIONONCORRUPTION};
	boolRegKey loadAddIns{L"LoadAddIns", true, false, IDS_REGKEY_LOADADDINS};
	boolRegKey forceOutlookMAPI{L"ForceOutlookMAPI", false, false, IDS_REGKEY_FORCEOUTLOOKMAPI};
	boolRegKey forceSystemMAPI{L"ForceSystemMAPI", false, false, IDS_REGKEY_FORCESYSTEMMAPI};
	boolRegKey hexDialogDiag{L"HexDialogDiag", false, false, IDS_REGKEY_HEXDIALOGDIAG};
	boolRegKey displayAboutDialog{L"DisplayAboutDialog", true, false, NULL};
	wstringRegKey propertyColumnOrder{L"PropertyColumnOrder", L"", false, NULL};

	// TODO: Can/should this be a vector?
	// Keep this in sync with REGKEYNAMES
	__RegKey* RegKeys[] = {
		&debugTag,
		&debugToFile,
		&debugFileName,
		&getPropNamesOnAllProps,
		&parseNamedProps,
		&throttleLevel,
		&hierExpandNotifications,
		&hierRootNotifs,
		&doSmartView,
		&onlyAdditionalProperties,
		&useRowDataForSinglePropList,
		&useGetPropList,
		&preferUnicodeProps,
		&cacheNamedProps,
		&allowDupeColumns,
		&doColumnNames,
		&editColumnsOnLoad,
		&forceMDBOnline,
		&forceMapiNoCache,
		&allowPersistCache,
		&useIMAPIProgress,
		&useMessageRaw,
		&suppressNotFound,
		&heapEnableTerminationOnCorruption,
		&loadAddIns,
		&forceOutlookMAPI,
		&forceSystemMAPI,
		&hexDialogDiag,
		&displayAboutDialog,
		&propertyColumnOrder,
	};

	void SetDefaults()
	{
		// Set some defaults to begin with:
		for (auto& regKey : RegKeys)
		{
			if (!regKey) continue;

			if (regKey->ulRegKeyType == regDWORD)
			{
				regKey->ulCurDWORD = regKey->ulDefDWORD;
			}
			else if (regKey->ulRegKeyType == regSTRING)
			{
				regKey->szCurSTRING = regKey->szDefSTRING;
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
				if (!regKey) continue;

				if (regKey->ulRegKeyType == regDWORD)
				{
					regKey->ulCurDWORD = ReadDWORDFromRegistry(hRootKey, regKey->szKeyName, regKey->ulCurDWORD);
				}
				else if (regKey->ulRegKeyType == regSTRING)
				{
					regKey->szCurSTRING = ReadStringFromRegistry(hRootKey, regKey->szKeyName, regKey->szCurSTRING);
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
			if (!regKey) continue;

			if (regKey->ulRegKeyType == regDWORD)
			{
				CommitDWORDIfNeeded(hRootKey, regKey->szKeyName, regKey->ulCurDWORD, regKey->ulDefDWORD);
			}
			else if (regKey->ulRegKeyType == regSTRING)
			{
				CommitStringIfNeeded(hRootKey, regKey->szKeyName, regKey->szCurSTRING, regKey->szDefSTRING);
			}
		}

		if (hRootKey) EC_W32_S(RegCloseKey(hRootKey));
	}
} // namespace registry