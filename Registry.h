// Registry.h : header file for the MFCMAPI application registry functions
//

#pragma once
#define RKEY_ROOT _T("SOFTWARE\\Microsoft\\MFCMAPI") // STRING_OK

enum __REGKEYTYPES {
	regDWORD,
	regSTRING
};

struct __RegKeys
{
	TCHAR*	szKeyName;
	ULONG	ulRegKeyType;
	ULONG	ulDefDWORD;
	ULONG	ulCurDWORD;
	TCHAR	szDefSTRING[MAX_PATH];
	TCHAR	szCurSTRING[MAX_PATH];
	UINT	uiOptionsPrompt;
};

//Registry key Names
enum REGKEYNAMES {
	regkeyDEBUG_TAG,
	regkeyDEBUG_TO_FILE,
	regkeyDEBUG_FILE_NAME,
	regkeyTHROTTLE_LEVEL,
	regkeyPARSED_NAMED_PROPS,
	regkeyHIER_NOTIFS,
	regkeyHIER_EXPAND_NOTIFS,
	regkeyHIER_ROOT_NOTIFS,
	regkeyHIER_NODE_LOAD_COUNT,
	regkeyDO_GETPROPS,
	regkeyALLOW_DUPE_COLUMNS,
	regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST,
	regkeyDO_COLUMN_NAMES,
	regkeyEDIT_COLUMNS_ON_LOAD,
	regkeyUSE_GETPROPLIST,
	regkeyGETPROPNAMES_ON_ALL_PROPS,
	regkeyMDB_ONLINE,
	regKeyMAPI_NO_CACHE,
	regkeyALLOW_PERSIST_CACHE,
	regkeyUSE_IMAPIPROGRESS,
	regkeyUSE_MESSAGERAW,
	regkeyDISPLAY_ABOUT_DIALOG,
	regkeyPROP_COLUMN_ORDER,
	NUMRegKeys
};

extern __RegKeys RegKeys[NUMRegKeys];

void SetDefaults();
void WriteToRegistry();
void ReadFromRegistry();

void ReadDWORDFromRegistry(HKEY hKey, LPCTSTR szValue, DWORD* lpdwVal);

HRESULT HrGetRegistryValue(
						   IN HKEY hKey, // the key.
						   IN LPCTSTR lpszValue, // value name in key.
						   OUT DWORD* lpType, // where to put type info.
						   OUT LPVOID* lppData); // where to put the data.