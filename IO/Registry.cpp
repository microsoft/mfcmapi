#include "stdafx.h"
#include <IO/Registry.h>

// Keep this in sync with REGKEYNAMES
__RegKeys RegKeys[] = {
#ifdef _DEBUG
	{ L"DebugTag",					regDWORD,regoptStringHex,	DBGAll		,0,L"",L"",false,	IDS_REGKEY_DEBUG_TAG }, // STRING_OK
#else
	{ L"DebugTag",					regDWORD,regoptStringHex,	DBGNoDebug	,0,L"",L"",false,	IDS_REGKEY_DEBUG_TAG }, // STRING_OK
#endif
	{ L"DebugToFile",				regDWORD,regoptCheck,		0			,0,L"",L"",false,	IDS_REGKEY_DEBUG_TO_FILE }, // STRING_OK
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
	{ L"DisplayAboutDialog",		regDWORD,regoptCheck,		true		,0,L"",L"",false,	NULL }, // STRING_OK
	{ L"PropertyColumnOrder",		regSTRING,regoptCheck,		0			,0,L"",L"",false,	NULL }, // STRING_OK
	// {KeyName,					keytype,opttype,			defaultDWORD,0,defaultString,NULL,bRefresh,IDS_REGKEY_*} // Regkey template
};

void SetDefaults()
{
	// Set some defaults to begin with:
	for (auto i = 0; i < NUMRegKeys; i++)
	{
		if (RegKeys[i].ulRegKeyType == regDWORD)
		{
			RegKeys[i].ulCurDWORD = RegKeys[i].ulDefDWORD;
		}
		else if (RegKeys[i].ulRegKeyType == regSTRING)
		{
			RegKeys[i].szCurSTRING = RegKeys[i].szDefSTRING;
		}
	}
}

// $--HrGetRegistryValue---------------------------------------------------------
// Get a registry value - allocating memory using new to hold it.
// -----------------------------------------------------------------------------
_Check_return_ HRESULT HrGetRegistryValue(
	_In_ HKEY hKey, // the key.
	_In_ const wstring& lpszValue, // value name in key.
	_Out_ DWORD* lpType, // where to put type info.
	_Out_ LPVOID* lppData) // where to put the data.
{
	auto hRes = S_OK;

	DebugPrint(DBGGeneric, L"HrGetRegistryValue(%ws)\n", lpszValue.c_str());

	*lppData = nullptr;
	DWORD cb = NULL;

	// Get its size
	WC_W32(RegQueryValueExW(
		hKey,
		lpszValue.c_str(),
		nullptr,
		lpType,
		nullptr,
		&cb));

	// only handle types we know about - all others are bad
	if (S_OK == hRes && cb &&
		(REG_SZ == *lpType || REG_DWORD == *lpType || REG_MULTI_SZ == *lpType))
	{
		*lppData = new BYTE[cb];

		if (*lppData)
		{
			// Get the current value
			EC_W32(RegQueryValueExW(
				hKey,
				lpszValue.c_str(),
				nullptr,
				lpType,
				static_cast<unsigned char*>(*lppData),
				&cb));

			if (FAILED(hRes))
			{
				delete[] * lppData;
				*lppData = nullptr;
			}
		}
	}
	else hRes = MAPI_E_INVALID_PARAMETER;

	return hRes;
}

// If the value is not set in the registry, return the default value
DWORD ReadDWORDFromRegistry(_In_ HKEY hKey, _In_ const wstring& szValue, _In_ DWORD dwDefaultVal)
{
	if (szValue.empty()) return dwDefaultVal;
	auto hRes = S_OK;
	DWORD dwKeyType = NULL;
	DWORD* lpValue = nullptr;
	auto ret = dwDefaultVal;

	WC_H(HrGetRegistryValue(
		hKey,
		szValue,
		&dwKeyType,
		reinterpret_cast<LPVOID*>(&lpValue)));
	if (hRes == S_OK && REG_DWORD == dwKeyType && lpValue)
	{
		ret = *lpValue;
	}

	delete[] lpValue;
	return ret;
}
// If the value is not set in the registry, return the default value
wstring ReadStringFromRegistry(_In_ HKEY hKey, _In_ const wstring& szValue, _In_ const wstring& szDefault)
{
	if (szValue.empty()) return szDefault;
	auto hRes = S_OK;
	DWORD dwKeyType = NULL;
	LPBYTE szBuf = nullptr;
	auto ret = szDefault;

	DebugPrint(DBGGeneric, L"ReadStringFromRegistry(%ws)\n", szValue.c_str());

	DWORD cb = NULL;

	// Get its size
	WC_W32(RegQueryValueExW(
		hKey,
		szValue.c_str(),
		nullptr,
		&dwKeyType,
		nullptr,
		&cb));

	if (S_OK == hRes && cb  && !(cb % 2) && REG_SZ == dwKeyType)
	{
		szBuf = new BYTE[cb];
		if (szBuf)
		{
			// Get the current value
			EC_W32(RegQueryValueExW(
				hKey,
				szValue.c_str(),
				nullptr,
				&dwKeyType,
				szBuf,
				&cb));

			if (FAILED(hRes))
			{
				delete[] szBuf;
				szBuf = nullptr;
			}
		}
	}

	if (S_OK == hRes && cb && !(cb % 2) && REG_SZ == dwKeyType && szBuf)
	{
		ret = wstring(LPWSTR(szBuf), cb / sizeof WCHAR);
	}

	delete[] szBuf;
	return ret;
}

void ReadFromRegistry()
{
	auto hRes = S_OK;
	HKEY hRootKey = nullptr;

	WC_W32(RegOpenKeyExW(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		NULL,
		KEY_READ,
		&hRootKey));

	// Now that we have a root key, go get our values
	if (hRootKey)
	{
		for (auto i = 0; i < NUMRegKeys; i++)
		{
			if (RegKeys[i].ulRegKeyType == regDWORD)
			{
				RegKeys[i].ulCurDWORD = ReadDWORDFromRegistry(
					hRootKey,
					RegKeys[i].szKeyName,
					RegKeys[i].ulCurDWORD);
			}
			else if (RegKeys[i].ulRegKeyType == regSTRING)
			{
				RegKeys[i].szCurSTRING = ReadStringFromRegistry(
					hRootKey,
					RegKeys[i].szKeyName,
					RegKeys[i].szCurSTRING);
			}
		}

		EC_W32(RegCloseKey(hRootKey));
	}

	SetDebugLevel(RegKeys[regkeyDEBUG_TAG].ulCurDWORD);
	DebugPrintVersion(DBGVersionBanner);
}

void WriteDWORDToRegistry(_In_ HKEY hKey, _In_ const wstring& szValueName, DWORD dwValue)
{
	auto hRes = S_OK;

	WC_W32(RegSetValueExW(
		hKey,
		szValueName.c_str(),
		NULL,
		REG_DWORD,
		reinterpret_cast<LPBYTE>(&dwValue),
		sizeof(DWORD)));
}

void CommitDWORDIfNeeded(_In_ HKEY hKey, _In_ const wstring& szValueName, DWORD dwValue, DWORD dwDefaultValue)
{
	auto hRes = S_OK;
	if (dwValue != dwDefaultValue)
	{
		WriteDWORDToRegistry(
			hKey,
			szValueName,
			dwValue);
	}
	else
	{
		WC_W32(RegDeleteValueW(hKey, szValueName.c_str()));
	}
}

void WriteStringToRegistry(_In_ HKEY hKey, _In_ const wstring& szValueName, _In_ const wstring& szValue)
{
	auto hRes = S_OK;

	// Reg needs bytes, so CB is correct here
	auto cbValue = szValue.length() * sizeof(WCHAR);
	cbValue += sizeof(WCHAR); // NULL terminator

	WC_W32(RegSetValueExW(
		hKey,
		szValueName.c_str(),
		NULL,
		REG_SZ,
		LPBYTE(szValue.c_str()),
		static_cast<DWORD>(cbValue)));
}

void CommitStringIfNeeded(_In_ HKEY hKey, _In_ const wstring& szValueName, _In_ const wstring& szValue, _In_ const wstring& szDefaultValue)
{
	auto hRes = S_OK;
	if (wstringToLower(szValue) != wstringToLower(szDefaultValue))
	{
		WriteStringToRegistry(
			hKey,
			szValueName,
			szValue);
	}
	else
	{
		WC_W32(RegDeleteValueW(hKey, szValueName.c_str()));
	}
}

_Check_return_ HKEY CreateRootKey()
{
	auto hRes = S_OK;
	HKEY hkSub = nullptr;

	// Try to open the root key before we do the work to create it
	WC_W32(RegOpenKeyExW(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		NULL,
		KEY_READ | KEY_WRITE,
		&hkSub));
	if (SUCCEEDED(hRes) && hkSub) return hkSub;

	hRes = S_OK;
	WC_W32(RegCreateKeyExW(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		0,
		nullptr,
		0,
		KEY_READ | KEY_WRITE,
		nullptr,
		&hkSub,
		nullptr));

	return hkSub;
}
void WriteToRegistry()
{
	auto hRes = S_OK;
	auto hRootKey = CreateRootKey();

	// Now that we have a root key, go set our values
	for (auto i = 0; i < NUMRegKeys; i++)
	{
		if (RegKeys[i].ulRegKeyType == regDWORD)
		{
			CommitDWORDIfNeeded(
				hRootKey,
				RegKeys[i].szKeyName,
				RegKeys[i].ulCurDWORD,
				RegKeys[i].ulDefDWORD);
		}
		else if (RegKeys[i].ulRegKeyType == regSTRING)
		{
			CommitStringIfNeeded(
				hRootKey,
				RegKeys[i].szKeyName,
				RegKeys[i].szCurSTRING,
				RegKeys[i].szDefSTRING);
		}
	}

	if (hRootKey) EC_W32(RegCloseKey(hRootKey));
}