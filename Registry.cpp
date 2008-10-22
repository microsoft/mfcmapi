#include "stdafx.h"
#include "registry.h"
#include <Aclapi.h>
#include "MAPIFunctions.h"

// Keep this in sync with REGKEYNAMES
__RegKeys RegKeys[] = {
#ifdef _DEBUG
	{_T("DebugTag"),					regDWORD,	DBGAll		,0,_T(""),_T(""),IDS_REGKEY_DEBUG_TAG}, // STRING_OK
#else
	{_T("DebugTag"),					regDWORD,	DBGNoDebug	,0,_T(""),_T(""),IDS_REGKEY_DEBUG_TAG }, // STRING_OK
#endif
	{_T("DebugToFile"),					regDWORD,	0			,0,_T(""),_T(""),IDS_REGKEY_DEBUG_TO_FILE}, // STRING_OK
	{_T("DebugFileName"),				regSTRING,	0			,0,_T("c:\\mfcmapi.log"),_T(""),IDS_REGKEY_DEBUG_FILE_NAME}, // STRING_OK
	{_T("ThrottleLevel"),				regDWORD,	0			,0,_T(""),_T(""),IDS_REGKEY_THROTTLE_LEVEL}, // STRING_OK
	{_T("ParseNamedProps"),				regDWORD,	true		,0,_T(""),_T(""),IDS_REGKEY_PARSED_NAMED_PROPS}, // STRING_OK
	{_T("HierNotifs"),					regDWORD,	true		,0,_T(""),_T(""),IDS_REGKEY_HIER_NOTIFS}, // STRING_OK
	{_T("HierExpandNotifs"),			regDWORD,	true		,0,_T(""),_T(""),IDS_REGKEY_HIER_EXPAND_NOTIFS}, // STRING_OK
	{_T("HierRootNotifs"),				regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_HIER_ROOT_NOTIFS}, // STRING_OK
	{_T("HierNodeLoadCount"),			regDWORD,	20			,0,_T(""),_T(""),IDS_REGKEY_HIER_NODE_LOAD_COUNT}, // STRING_OK
	{_T("DoGetProps"),					regDWORD,	true		,0,_T(""),_T(""),IDS_REGKEY_DO_GETPROPS}, // STRING_OK
	{_T("AllowDupeColumns"),			regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_ALLOW_DUPE_COLUMNS}, // STRING_OK
	{_T("UseRowDataForSinglePropList"),	regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_USE_ROW_DATA_FOR_SINGLEPROPLIST}, // STRING_OK
	{_T("DoColumnNames"),				regDWORD,	true		,0,_T(""),_T(""),IDS_REGKEY_DO_COLUMN_NAMES}, // STRING_OK
	{_T("EditColumnsOnLoad"),			regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_EDIT_COLUMNS_ON_LOAD}, // STRING_OK
	{_T("UseGetPropList"),				regDWORD,	true		,0,_T(""),_T(""),IDS_REGKEY_USE_GETPROPLIST}, // STRING_OK
	{_T("GetPropNamesOnAllProps"),		regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_GETPROPNAMES_ON_ALL_PROPS}, // STRING_OK
	{_T("ForceMDBOnline"),				regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_MDB_ONLINE}, // STRING_OK
	{_T("ForceMapiNoCache"),			regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_MAPI_NO_CACHE}, // STRING_OK
	{_T("AllowPersistCache"),			regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_ALLOW_PERSIST_CACHE }, // STRING_OK
	{_T("UseIMAPIProgress"),			regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_USE_IMAPIPROGRESS}, // STRING_OK
	{_T("UseMessageRaw"),				regDWORD,	false		,0,_T(""),_T(""),IDS_REGKEY_USE_MESSAGERAW}, // STRING_OK
	{_T("DisplayAboutDialog"),			regDWORD,	true		,0,_T(""),_T(""),NULL}, // STRING_OK
	{_T("PropertyColumnOrder"),			regSTRING,	0			,0,_T(""),_T(""),NULL}, // STRING_OK
//	{KeyName,							keytype,	defaultDWORD,0,defaultString,NULL,IDS_REGKEY_*} // Regkey template
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
				CCH(RegKeys[i].szCurSTRING),
				RegKeys[i].szDefSTRING));
		}

	}
}

// $--HrGetRegistryValue---------------------------------------------------------
// Get a registry value - allocating memory using new to hold it.
// -----------------------------------------------------------------------------
HRESULT HrGetRegistryValue(
						   IN HKEY hKey, // the key.
						   IN LPCTSTR lpszValue, // value name in key.
						   OUT DWORD* lpType, // where to put type info.
						   OUT LPVOID* lppData) // where to put the data.
{
	HRESULT hRes = S_OK;

	DebugPrint(DBGGeneric,_T("HrGetRegistryValue(%s)\n"),lpszValue);

	*lppData = NULL;
	DWORD cb = NULL;

	// Get its size
	WC_W32(RegQueryValueEx(
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
			EC_W32(RegQueryValueEx(
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
}

// If the value is not set in the registry, do not alter the passed in DWORD
void ReadDWORDFromRegistry(HKEY hKey, LPCTSTR szValue, DWORD* lpdwVal)
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
}

// If the value is not set in the registry, do not alter the passed in string
void ReadStringFromRegistry(HKEY hKey, LPCTSTR szValue, LPTSTR szDest, ULONG cchDestLen)
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
}

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
					CCH(RegKeys[i].szCurSTRING));
				if (RegKeys[i].szCurSTRING[0] == _T('\0'))
				{
					EC_H(StringCchCopy(
						RegKeys[i].szCurSTRING,
						CCH(RegKeys[i].szCurSTRING),
						RegKeys[i].szDefSTRING));
				}
			}

		}

		EC_W32(RegCloseKey(hRootKey));
	}

	SetDebugLevel(RegKeys[regkeyDEBUG_TAG].ulCurDWORD);
}

void WriteDWORDToRegistry(HKEY hKey, LPCTSTR szValueName, DWORD dwValue)
{
	HRESULT hRes = S_OK;

	WC_W32(RegSetValueEx(
		hKey,
		szValueName,
		NULL,
		REG_DWORD,
		(LPBYTE) &dwValue,
		sizeof(DWORD)));
}

void CommitDWORDIfNeeded(HKEY hKey, LPCTSTR szValueName, DWORD dwValue, DWORD dwDefaultValue)
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
}

void WriteStringToRegistry(HKEY hKey, LPCTSTR szValueName, LPCTSTR szValue)
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
}

void CommitStringIfNeeded(HKEY hKey, LPCTSTR szValueName, LPCTSTR szValue, LPCTSTR szDefaultValue)
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
}

HKEY CreateRootKey()
{
	HRESULT hRes = S_OK;
    PSID pEveryoneSID = NULL;
    PACL pACL = NULL;
    PSECURITY_DESCRIPTOR pSD = NULL;
    EXPLICIT_ACCESS ea[2];
    SID_IDENTIFIER_AUTHORITY SIDAuthWorld = SECURITY_WORLD_SID_AUTHORITY;
    SECURITY_ATTRIBUTES sa = {0};
    HKEY hkSub = NULL;

    // Create a well-known SID for the Everyone group.
    EC_B(AllocateAndInitializeSid(
		&SIDAuthWorld,
		1,
		SECURITY_WORLD_RID,
		0, 0, 0, 0, 0, 0, 0,
		&pEveryoneSID));

    // Initialize an EXPLICIT_ACCESS structure for an ACE.
    // The ACE will allow Everyone full control to the key.
    ZeroMemory(ea, sizeof(EXPLICIT_ACCESS));
    ea[0].grfAccessPermissions = KEY_ALL_ACCESS;
    ea[0].grfAccessMode = SET_ACCESS;
    ea[0].grfInheritance= SUB_CONTAINERS_AND_OBJECTS_INHERIT;
    ea[0].Trustee.TrusteeForm = TRUSTEE_IS_SID;
    ea[0].Trustee.TrusteeType = TRUSTEE_IS_WELL_KNOWN_GROUP;
    ea[0].Trustee.ptstrName  = (LPTSTR) pEveryoneSID;

    // Create a new ACL that contains the new ACEs.
    EC_W32(SetEntriesInAcl(1, ea, NULL, &pACL));

    // Initialize a security descriptor.
    EC_D(pSD, LocalAlloc(LPTR,
		SECURITY_DESCRIPTOR_MIN_LENGTH));

    EC_B(InitializeSecurityDescriptor(pSD, SECURITY_DESCRIPTOR_REVISION));

    // Add the ACL to the security descriptor.
    EC_B(SetSecurityDescriptorDacl(pSD,
		TRUE,     // bDaclPresent flag
		pACL,
		FALSE));   // not a default DACL

    // Initialize a security attributes structure.
    sa.nLength = sizeof (SECURITY_ATTRIBUTES);
    sa.lpSecurityDescriptor = pSD;
    sa.bInheritHandle = FALSE;

    // Use the security attributes to set the security descriptor
    // when you create a key.
    WC_W32(RegCreateKeyEx(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		0,
		_T(""),
		0,
		KEY_READ | KEY_WRITE,
		&sa,
		&hkSub,
		NULL));

    if (pEveryoneSID) FreeSid(pEveryoneSID);
    LocalFree(pACL);
    LocalFree(pSD);

	return hkSub;
}

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
}
