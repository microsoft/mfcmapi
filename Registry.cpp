#include "stdafx.h"
#include "registry.h"

// Keep this in sync with REGKEYNAMES
__RegKeys RegKeys[] = {
#ifdef _DEBUG
	{_T("DebugTag"),					regDWORD,regoptStringHex,	DBGAll		,0,_T(""),_T(""),false,	IDS_REGKEY_DEBUG_TAG}, // STRING_OK
#else
	{_T("DebugTag"),					regDWORD,regoptStringHex,	DBGNoDebug	,0,_T(""),_T(""),false,	IDS_REGKEY_DEBUG_TAG }, // STRING_OK
#endif
	{_T("DebugToFile"),					regDWORD,regoptCheck,		0			,0,_T(""),_T(""),false,	IDS_REGKEY_DEBUG_TO_FILE}, // STRING_OK
	{_T("DebugFileName"),				regSTRING,regoptString,		0			,0,_T("c:\\mfcmapi.log"),_T(""),false,	IDS_REGKEY_DEBUG_FILE_NAME}, // STRING_OK
	{_T("ParseNamedProps"),				regDWORD,regoptCheck,		true		,0,_T(""),_T(""),true,	IDS_REGKEY_PARSED_NAMED_PROPS}, // STRING_OK
	{_T("GetPropNamesOnAllProps"),		regDWORD,regoptCheck,		false		,0,_T(""),_T(""),true,	IDS_REGKEY_GETPROPNAMES_ON_ALL_PROPS}, // STRING_OK
	{_T("ThrottleLevel"),				regDWORD,regoptStringDec,	0			,0,_T(""),_T(""),false,	IDS_REGKEY_THROTTLE_LEVEL}, // STRING_OK
	{_T("HierExpandNotifications"),		regDWORD,regoptCheck,		true		,0,_T(""),_T(""),false,	IDS_REGKEY_HIER_EXPAND_NOTIFS}, // STRING_OK
	{_T("HierRootNotifs"),				regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_HIER_ROOT_NOTIFS}, // STRING_OK
	{_T("DoSmartView"),					regDWORD,regoptCheck,		true		,0,_T(""),_T(""),true,	IDS_REGKEY_DO_SMART_VIEW}, // STRING_OK
	{_T("DoGetProps"),					regDWORD,regoptCheck,		true		,0,_T(""),_T(""),true,	IDS_REGKEY_DO_GETPROPS}, // STRING_OK
	{_T("UseGetPropList"),				regDWORD,regoptCheck,		true		,0,_T(""),_T(""),true,	IDS_REGKEY_USE_GETPROPLIST}, // STRING_OK
	{_T("CacheNamedProps"),				regDWORD,regoptCheck,		true		,0,_T(""),_T(""),false,	IDS_REGKEY_CACHE_NAMED_PROPS}, // STRING_OK
	{_T("AllowDupeColumns"),			regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_ALLOW_DUPE_COLUMNS}, // STRING_OK
	{_T("UseRowDataForSinglePropList"),	regDWORD,regoptCheck,		false		,0,_T(""),_T(""),true,	IDS_REGKEY_USE_ROW_DATA_FOR_SINGLEPROPLIST}, // STRING_OK
	{_T("DoColumnNames"),				regDWORD,regoptCheck,		true		,0,_T(""),_T(""),false,	IDS_REGKEY_DO_COLUMN_NAMES}, // STRING_OK
	{_T("EditColumnsOnLoad"),			regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_EDIT_COLUMNS_ON_LOAD}, // STRING_OK
	{_T("ForceMDBOnline"),				regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_MDB_ONLINE}, // STRING_OK
	{_T("ForceMapiNoCache"),			regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_MAPI_NO_CACHE}, // STRING_OK
	{_T("AllowPersistCache"),			regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_ALLOW_PERSIST_CACHE }, // STRING_OK
	{_T("UseIMAPIProgress"),			regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_USE_IMAPIPROGRESS}, // STRING_OK
	{_T("UseMessageRaw"),				regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_USE_MESSAGERAW}, // STRING_OK
	{_T("HeapEnableTerminationOnCorruption"),regDWORD,regoptCheck,	true		,0,_T(""),_T(""),false,	IDS_REGKEY_HEAPENABLETERMINATIONONCORRUPTION}, // STRING_OK
	{_T("LoadAddIns"),					regDWORD,regoptCheck,		true		,0,_T(""),_T(""),false,	IDS_REGKEY_LOADADDINS}, // STRING_OK
	{_T("ForceOutlookMAPI"),			regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_FORCEOUTLOOKMAPI}, // STRING_OK
	{_T("ForceSystemMAPI"),				regDWORD,regoptCheck,		false		,0,_T(""),_T(""),false,	IDS_REGKEY_FORCESYSTEMMAPI}, // STRING_OK
	{_T("DisplayAboutDialog"),			regDWORD,regoptCheck,		true		,0,_T(""),_T(""),false,	NULL}, // STRING_OK
	{_T("PropertyColumnOrder"),			regSTRING,regoptCheck,		0			,0,_T(""),_T(""),false,	NULL}, // STRING_OK
	// {KeyName,						keytype,opttype,			defaultDWORD,0,defaultString,NULL,bRefresh,IDS_REGKEY_*} // Regkey template
};

void SetDefaults()
{
	HRESULT hRes = S_OK;
	// Set some defaults to begin with:
	int i = 0;

	for (i = 0 ; i < NUMRegKeys ; i++)
	{
		if (RegKeys[i].ulRegKeyType == regDWORD)
		{
			RegKeys[i].ulCurDWORD = RegKeys[i].ulDefDWORD;
		}
		else if (RegKeys[i].ulRegKeyType == regSTRING)
		{
			EC_H(StringCchCopy(
				RegKeys[i].szCurSTRING,
				_countof(RegKeys[i].szCurSTRING),
				RegKeys[i].szDefSTRING));
		}
	}
} // SetDefaults

// $--HrGetRegistryValueW---------------------------------------------------------
// Get a registry value - allocating memory using new to hold it.
// -----------------------------------------------------------------------------
_Check_return_ HRESULT HrGetRegistryValueW(
	_In_ HKEY hKey, // the key.
	_In_z_ LPCWSTR lpszValue, // value name in key.
	_Out_ DWORD* lpType, // where to put type info.
	_Out_ LPVOID* lppData) // where to put the data.
{
	HRESULT hRes = S_OK;

	DebugPrint(DBGGeneric,_T("HrGetRegistryValue(%ws)\n"),lpszValue);

	*lppData = NULL;
	DWORD cb = NULL;

	// Get its size
	WC_W32(RegQueryValueExW(
		hKey,
		lpszValue,
		NULL,
		lpType,
		NULL,
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
				lpszValue,
				NULL,
				lpType,
				(unsigned char*)*lppData,
				&cb));

			if (FAILED(hRes))
			{
				delete[] *lppData;
				*lppData = NULL;
			}
		}
	}
	else hRes = MAPI_E_INVALID_PARAMETER;

	return hRes;
} // HrGetRegistryValueW

// $--HrGetRegistryValueA---------------------------------------------------------
// Get a registry value - allocating memory using new to hold it.
// -----------------------------------------------------------------------------
_Check_return_ HRESULT HrGetRegistryValueA(
	_In_ HKEY hKey, // the key.
	_In_z_ LPCSTR lpszValue, // value name in key.
	_Out_ DWORD* lpType, // where to put type info.
	_Out_ LPVOID* lppData) // where to put the data.
{
	HRESULT hRes = S_OK;

	DebugPrint(DBGGeneric,_T("HrGetRegistryValueA(%hs)\n"),lpszValue);

	*lppData = NULL;
	DWORD cb = NULL;

	// Get its size
	WC_W32(RegQueryValueExA(
		hKey,
		lpszValue,
		NULL,
		lpType,
		NULL,
		&cb));

	// only handle types we know about - all others are bad
	if (S_OK == hRes && cb &&
		(REG_SZ == *lpType || REG_DWORD == *lpType || REG_MULTI_SZ == *lpType))
	{
		*lppData = new BYTE[cb];

		if (*lppData)
		{
			// Get the current value
			EC_W32(RegQueryValueExA(
				hKey,
				lpszValue,
				NULL,
				lpType,
				(unsigned char*)*lppData,
				&cb));

			if (FAILED(hRes))
			{
				delete[] *lppData;
				*lppData = NULL;
			}
		}
	}
	else if (SUCCEEDED(hRes))
	{
		hRes = MAPI_E_INVALID_PARAMETER;
	}

	return hRes;
} // HrGetRegistryValueA

// If the value is not set in the registry, do not alter the passed in DWORD
void ReadDWORDFromRegistry(_In_ HKEY hKey, _In_z_ LPCTSTR szValue, _Out_ DWORD* lpdwVal)
{
	HRESULT hRes = S_OK;

	if (!szValue || !lpdwVal) return;
	DWORD dwKeyType = NULL;
	DWORD* lpValue = NULL;

	WC_H(HrGetRegistryValue(
		hKey,
		szValue,
		&dwKeyType,
		(LPVOID*) &lpValue));
	if (hRes == S_OK && REG_DWORD == dwKeyType && lpValue)
	{
		*lpdwVal = *lpValue;
	}

	delete[] lpValue;
} // ReadDWORDFromRegistry

// If the value is not set in the registry, do not alter the passed in string
void ReadStringFromRegistry(_In_ HKEY hKey, _In_z_ LPCTSTR szValue, _In_z_ LPTSTR szDest, ULONG cchDestLen)
{
	HRESULT hRes = S_OK;

	if (!szValue || !szDest || (cchDestLen < 1)) return;

	DWORD dwKeyType = NULL;
	LPTSTR szBuf = NULL;

	WC_H(HrGetRegistryValue(
		hKey,
		szValue,
		&dwKeyType,
		(LPVOID*) &szBuf));
	if (hRes == S_OK && REG_SZ == dwKeyType && szBuf)
	{
		WC_H(StringCchCopy(szDest,cchDestLen,szBuf));
	}

	delete[] szBuf;
} // ReadStringFromRegistry

void ReadFromRegistry()
{
	HRESULT hRes = S_OK;
	HKEY hRootKey = NULL;

	WC_W32(RegOpenKeyEx(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		NULL,
		KEY_READ,
		&hRootKey));

	// Now that we have a root key, go get our values
	if (hRootKey)
	{
		int i = 0;

		for (i = 0 ; i < NUMRegKeys ; i++)
		{
			if (RegKeys[i].ulRegKeyType == regDWORD)
			{
				ReadDWORDFromRegistry(
					hRootKey,
					RegKeys[i].szKeyName,
					&RegKeys[i].ulCurDWORD);
			}
			else if (RegKeys[i].ulRegKeyType == regSTRING)
			{
				ReadStringFromRegistry(
					hRootKey,
					RegKeys[i].szKeyName,
					RegKeys[i].szCurSTRING,
					_countof(RegKeys[i].szCurSTRING));
				if (RegKeys[i].szCurSTRING[0] == _T('\0'))
				{
					EC_H(StringCchCopy(
						RegKeys[i].szCurSTRING,
						_countof(RegKeys[i].szCurSTRING),
						RegKeys[i].szDefSTRING));
				}
			}
		}

		EC_W32(RegCloseKey(hRootKey));
	}

	SetDebugLevel(RegKeys[regkeyDEBUG_TAG].ulCurDWORD);
	DebugPrintVersion(DBGVersionBanner);
} // ReadFromRegistry

void WriteDWORDToRegistry(_In_ HKEY hKey, _In_z_ LPCTSTR szValueName, DWORD dwValue)
{
	HRESULT hRes = S_OK;

	WC_W32(RegSetValueEx(
		hKey,
		szValueName,
		NULL,
		REG_DWORD,
		(LPBYTE) &dwValue,
		sizeof(DWORD)));
} // WriteDWORDToRegistry

void CommitDWORDIfNeeded(_In_ HKEY hKey, _In_z_ LPCTSTR szValueName, DWORD dwValue, DWORD dwDefaultValue)
{
	HRESULT hRes = S_OK;
	if (dwValue != dwDefaultValue)
	{
		WriteDWORDToRegistry(
			hKey,
			szValueName,
			dwValue);
	}
	else
	{
		WC_W32(RegDeleteValue(hKey,szValueName));
		hRes = S_OK;
	}
} // CommitDWORDIfNeeded

void WriteStringToRegistry(_In_ HKEY hKey, _In_z_ LPCTSTR szValueName, _In_z_ LPCTSTR szValue)
{
	HRESULT hRes = S_OK;
	size_t cbValue = 0;

	// Reg needs bytes, so CB is correct here
	EC_H(StringCbLength(szValue,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbValue));
	cbValue += sizeof(TCHAR); // NULL terminator

	WC_W32(RegSetValueEx(
		hKey,
		szValueName,
		NULL,
		REG_SZ,
		(LPBYTE) szValue,
		(DWORD) cbValue));
} // WriteStringToRegistry

void CommitStringIfNeeded(_In_ HKEY hKey, _In_z_ LPCTSTR szValueName, _In_z_ LPCTSTR szValue, _In_z_ LPCTSTR szDefaultValue)
{
	HRESULT hRes = S_OK;
	if (0 != lstrcmpi(szValue,szDefaultValue))
	{
		WriteStringToRegistry(
			hKey,
			szValueName,
			szValue);
	}
	else
	{
		WC_W32(RegDeleteValue(hKey,szValueName));
		hRes = S_OK;
	}
} // CommitStringIfNeeded

_Check_return_ HKEY CreateRootKey()
{
	HRESULT hRes = S_OK;
	HKEY hkSub = NULL;

	// Try to open the root key before we do the work to create it
	WC_W32(RegOpenKeyEx(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		NULL,
		KEY_READ | KEY_WRITE,
		&hkSub));
	if (SUCCEEDED(hRes) && hkSub) return hkSub;

	hRes = S_OK;
	WC_W32(RegCreateKeyEx(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		0,
		NULL,
		0,
		KEY_READ | KEY_WRITE,
		NULL,
		&hkSub,
		NULL));

	return hkSub;
} // CreateRootKey

void WriteToRegistry()
{
	HRESULT hRes = S_OK;
	HKEY hRootKey = NULL;

	hRootKey = CreateRootKey();

	// Now that we have a root key, go set our values

	int i = 0;

	for (i = 0 ; i < NUMRegKeys ; i++)
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
} // WriteToRegistry