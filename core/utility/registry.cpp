#include <core/stdafx.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <core/utility/import.h>
#include <mapistub/library/stubutils.h>

namespace registry
{
	// Registry settings
	// Creating a new reg key:
	// 1 - Define an accessor object using boolRegKey, wstringRegKey or dwordRegKey
	// 2 - Add pointer to the object to RegKeys vector
	// 3 - Add an extern for the object to registry.h
	// 4 - If the setting should show in options, ensure it has a unique options prompt value
	// Note: Accessor objects can be used in code as their underlying type, though some care may be needed with casting
#ifdef _DEBUG
	dwordRegKey debugTag{
		L"DebugTag",
		regOptionType::stringHex,
		static_cast<DWORD>(output::dbgLevel::All),
		false,
		IDS_REGKEY_DEBUG_TAG};
#else
	dwordRegKey debugTag{
		L"DebugTag",
		regOptionType::stringHex,
		static_cast<DWORD>(output::dbgLevel::NoDebug),
		false,
		IDS_REGKEY_DEBUG_TAG};
#endif
	boolRegKey debugToFile{L"DebugToFile", false, false, IDS_REGKEY_DEBUG_TO_FILE};
	wstringRegKey debugFileName{L"DebugFileName", L"c:\\mfcmapi.log", false, IDS_REGKEY_DEBUG_FILE_NAME};
	boolRegKey getPropNamesOnAllProps{L"GetPropNamesOnAllProps", false, true, IDS_REGKEY_GETPROPNAMES_ON_ALL_PROPS};
	boolRegKey parseNamedProps{L"ParseNamedProps", true, true, IDS_REGKEY_PARSED_NAMED_PROPS};
	dwordRegKey throttleLevel{L"ThrottleLevel", regOptionType::stringDec, 0, false, IDS_REGKEY_THROTTLE_LEVEL};
	boolRegKey hierExpandNotifications{L"HierExpandNotifications", true, false, IDS_REGKEY_HIER_EXPAND_NOTIFS};
	boolRegKey hierRootNotifs{L"HierRootNotifs", false, false, IDS_REGKEY_HIER_ROOT_NOTIFS};
	boolRegKey doSmartView{L"DoSmartView", true, true, IDS_REGKEY_DO_SMART_VIEW};
	boolRegKey onlyAdditionalProperties{L"OnlyAdditionalProperties", false, true, IDS_REGKEY_ONLYADDITIONALPROPERTIES};
	boolRegKey useRowDataForSinglePropList{
		L"UseRowDataForSinglePropList",
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
	boolRegKey heapEnableTerminationOnCorruption{
		L"HeapEnableTerminationOnCorruption",
		true,
		false,
		IDS_REGKEY_HEAPENABLETERMINATIONONCORRUPTION};
	boolRegKey loadAddIns{L"LoadAddIns", true, false, IDS_REGKEY_LOADADDINS};
	boolRegKey forceOutlookMAPI{L"ForceOutlookMAPI", false, false, IDS_REGKEY_FORCEOUTLOOKMAPI};
	boolRegKey forceSystemMAPI{L"ForceSystemMAPI", false, false, IDS_REGKEY_FORCESYSTEMMAPI};
	boolRegKey preferOlmapi32{L"PreferOlmapi32", true, false, IDS_REGKEY_PREFER_OLMAPI32};
	boolRegKey uiDiag{L"UIDiag", false, false, IDS_REGKEY_UIDIAG};
	boolRegKey displayAboutDialog{L"DisplayAboutDialog", true, false, NULL};
	wstringRegKey propertyColumnOrder{L"PropertyColumnOrder", L"", false, NULL};
	dwordRegKey namedPropBatchSize{L"NamedPropBatchSize", regOptionType::stringDec, 400, false, NULL};

	std::vector<__RegKey*> RegKeys = {
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
		&preferOlmapi32,
		&uiDiag,
		&displayAboutDialog,
		&propertyColumnOrder,
		&namedPropBatchSize};

	// If the value is not set in the registry, return the default value
	DWORD ReadDWORDFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ const DWORD dwDefaultVal)
	{
		if (szValue.empty()) return dwDefaultVal;
		output::DebugPrint(output::dbgLevel::Generic, L"ReadDWORDFromRegistry(%ws)\n", szValue.c_str());

		// Get its size
		DWORD cb{};
		DWORD dwKeyType{};
		auto hRes = WC_W32(RegQueryValueExW(hKey, szValue.c_str(), nullptr, &dwKeyType, nullptr, &cb));

		if (hRes == S_OK && cb && dwKeyType == REG_DWORD)
		{
			auto szBuf = std::vector<BYTE>(cb);

			// Get the current value
			hRes = EC_W32(RegQueryValueExW(hKey, szValue.c_str(), nullptr, &dwKeyType, szBuf.data(), &cb));
			if (hRes == S_OK && dwKeyType == REG_DWORD)
			{
				return *reinterpret_cast<DWORD*>(szBuf.data());
			}
		}

		return dwDefaultVal;
	}

	// If the value is not set in the registry, return the default value
	std::wstring
	ReadStringFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ const std::wstring& szDefault)
	{
		if (szValue.empty()) return szDefault;
		output::DebugPrint(output::dbgLevel::Generic, L"ReadStringFromRegistry(%ws)\n", szValue.c_str());

		// Get its size
		DWORD cb{};
		DWORD dwKeyType{};
		auto hRes = WC_W32(RegQueryValueExW(hKey, szValue.c_str(), nullptr, &dwKeyType, nullptr, &cb));

		if (hRes == S_OK && cb && !(cb % 2) && dwKeyType == REG_SZ)
		{
			auto szBuf = std::wstring(cb / 2, '\0');
			// Get the current value
			hRes = EC_W32(RegQueryValueExW(
				hKey,
				szValue.c_str(),
				nullptr,
				&dwKeyType,
				reinterpret_cast<LPBYTE>(const_cast<wchar_t*>(szBuf.data())),
				&cb));
			if (hRes == S_OK && cb && !(cb % 2) && dwKeyType == REG_SZ)
			{
				return szBuf;
			}
		}

		return szDefault;
	}

	std::vector<BYTE> ReadBinFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ const bool bSecure)
	{
		if (szValue.empty()) return {};
		output::DebugPrint(output::dbgLevel::Generic, L"ReadBinFromRegistry(%ws)\n", szValue.c_str());

		// Get its size
		DWORD cb{};
		DWORD dwKeyType{};
		auto hRes = WC_W32(RegQueryValueExW(hKey, szValue.c_str(), nullptr, &dwKeyType, nullptr, &cb));

		if (hRes == S_OK && cb && dwKeyType == REG_BINARY)
		{
			auto bin = std::vector<BYTE>(cb);
			// Get the current value
			hRes = EC_W32(
				RegQueryValueExW(hKey, szValue.c_str(), nullptr, &dwKeyType, const_cast<LPBYTE>(bin.data()), &cb));
			if (hRes == S_OK && cb && dwKeyType == REG_BINARY)
			{
				if (bSecure)
				{
					auto DataIn = DATA_BLOB{static_cast<DWORD>(bin.size()), bin.data()};
					auto DataOut = DATA_BLOB{};
					if (import::pfnCryptUnprotectData(&DataIn, nullptr, nullptr, nullptr, nullptr, 0, &DataOut))
					{
						bin = std::vector<BYTE>(DataOut.pbData, DataOut.pbData + DataOut.cbData);
					}

					LocalFree(DataOut.pbData);
				}

				return bin;
			}
		}

		return {};
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

				if (regKey->ulRegKeyType == regKeyType::dword)
				{
					regKey->ulCurDWORD = ReadDWORDFromRegistry(hRootKey, regKey->szKeyName, regKey->ulCurDWORD);
				}
				else if (regKey->ulRegKeyType == regKeyType::string)
				{
					regKey->szCurSTRING = ReadStringFromRegistry(hRootKey, regKey->szKeyName, regKey->szCurSTRING);
				}
			}

			EC_W32_S(RegCloseKey(hRootKey));
		}

		output::outputVersion(output::dbgLevel::VersionBanner, nullptr);
	}

	void WriteDWORDToRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValueName, DWORD dwValue)
	{
		WC_W32_S(RegSetValueExW(
			hKey, szValueName.c_str(), NULL, REG_DWORD, reinterpret_cast<LPBYTE>(&dwValue), sizeof(DWORD)));
	}

	void CommitDWORDIfNeeded(
		_In_ HKEY hKey,
		_In_ const std::wstring& szValueName,
		const DWORD dwValue,
		const DWORD dwDefaultValue)
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
		if (!strings::compareInsensitive(szValue, szDefaultValue))
		{
			WriteStringToRegistry(hKey, szValueName, szValue);
		}
		else
		{
			WC_W32_S(RegDeleteValueW(hKey, szValueName.c_str()));
		}
	}

	void WriteBinToRegistry(
		_In_ HKEY hKey,
		_In_ const std::wstring& szValueName,
		_In_ const std::vector<BYTE>& binValue,
		_In_ const bool bSecure)
	{
		const DWORD cbValue = static_cast<DWORD>(binValue.size());
		if (bSecure)
		{
			auto DataIn = DATA_BLOB{cbValue, const_cast<LPBYTE>(binValue.data())};
			auto DataOut = DATA_BLOB{};
			if (import::pfnCryptProtectData(&DataIn, nullptr, nullptr, nullptr, nullptr, 0, &DataOut))
			{
				WC_W32_S(RegSetValueExW(hKey, szValueName.c_str(), NULL, REG_BINARY, DataOut.pbData, DataOut.cbData));
			}

			LocalFree(DataOut.pbData);
		}
		else
		{
			WC_W32_S(RegSetValueExW(hKey, szValueName.c_str(), NULL, REG_BINARY, binValue.data(), cbValue));
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

			if (regKey->ulRegKeyType == regKeyType::dword)
			{
				CommitDWORDIfNeeded(hRootKey, regKey->szKeyName, regKey->ulCurDWORD, regKey->ulDefDWORD);
			}
			else if (regKey->ulRegKeyType == regKeyType::string)
			{
				CommitStringIfNeeded(hRootKey, regKey->szKeyName, regKey->szCurSTRING, regKey->szDefSTRING);
			}
		}

		if (hRootKey) EC_W32_S(RegCloseKey(hRootKey));
	}

	void PushOptionsToStub()
	{
		mapistub::ForceOutlookMAPI(forceOutlookMAPI);
		mapistub::ForceSystemMAPI(forceSystemMAPI);
		mapistub::PreferOlmapi32(preferOlmapi32);
	}
} // namespace registry