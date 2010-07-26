// Registry.h : header file for the MFCMAPI application registry functions
//

#pragma once
#define RKEY_ROOT _T("SOFTWARE\\Microsoft\\MFCMAPI") // STRING_OK

enum __REGKEYTYPES {
	regDWORD,
	regSTRING
};

enum __REGOPTIONTYPE {
	regoptCheck,
	regoptString,
	regoptStringHex,
	regoptStringDec
};

struct __RegKeys
{
	TCHAR*	szKeyName;
	ULONG	ulRegKeyType;
	ULONG	ulRegOptType;
	ULONG	ulDefDWORD;
	ULONG	ulCurDWORD;
	TCHAR	szDefSTRING[MAX_PATH];
	TCHAR	szCurSTRING[MAX_PATH];
	BOOL	bRefresh;
	UINT	uiOptionsPrompt;
};

// Registry key Names
enum REGKEYNAMES {
	regkeyDEBUG_TAG,
	regkeyDEBUG_TO_FILE,
	regkeyDEBUG_FILE_NAME,
	regkeyPARSED_NAMED_PROPS,
	regkeyGETPROPNAMES_ON_ALL_PROPS,
	regkeyTHROTTLE_LEVEL,
	regkeyHIER_NOTIFS,
	regkeyHIER_EXPAND_NOTIFS,
	regkeyHIER_ROOT_NOTIFS,
	regkeyHIER_NODE_LOAD_COUNT,
	regkeyDO_SMART_VIEW,
	regkeyDO_GETPROPS,
	regkeyUSE_GETPROPLIST,
	regkeyCACHE_NAME_DPROPS,
	regkeyALLOW_DUPE_COLUMNS,
	regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST,
	regkeyDO_COLUMN_NAMES,
	regkeyEDIT_COLUMNS_ON_LOAD,
	regkeyMDB_ONLINE,
	regKeyMAPI_NO_CACHE,
	regkeyALLOW_PERSIST_CACHE,
	regkeyUSE_IMAPIPROGRESS,
	regkeyUSE_MESSAGERAW,
	regkeyHEAPENABLETERMINATIONONCORRUPTION,
	regkeyDISPLAY_ABOUT_DIALOG,
	regkeyPROP_COLUMN_ORDER,
	NUMRegKeys
};

#define NumRegOptionKeys (NUMRegKeys-2)

extern __RegKeys RegKeys[NUMRegKeys];

void SetDefaults();
void WriteToRegistry();
void ReadFromRegistry();

void ReadDWORDFromRegistry(_In_ HKEY hKey, _In_z_ LPCTSTR szValue, _Out_ DWORD* lpdwVal);

_Check_return_ HRESULT HrGetRegistryValueW(
	_In_ HKEY hKey, // the key.
	_In_z_ LPCTSTR lpszValue, // value name in key.
	_Out_ DWORD* lpType, // where to put type info.
	_Out_ LPVOID* lppData); // where to put the data.
_Check_return_ HRESULT HrGetRegistryValueA(
	_In_ HKEY hKey, // the key.
	_In_z_ LPCSTR lpszValue, // value name in key.
	_Out_ DWORD* lpType, // where to put type info.
	_Out_ LPVOID* lppData); // where to put the data.
#ifdef UNICODE
#define HrGetRegistryValue HrGetRegistryValueW
#else
#define HrGetRegistryValue HrGetRegistryValueA
#endif