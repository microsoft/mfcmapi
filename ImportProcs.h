// ImportProcs.h : header file for loading imports from DLLs
#pragma once

extern LPEDITSECURITY pfnEditSecurity;
extern LPSTGCREATESTORAGEEX pfnStgCreateStorageEx;
extern LPOPENTHEMEDATA pfnOpenThemeData;
extern LPCLOSETHEMEDATA pfnCloseThemeData;
extern LPGETTHEMEMARGINS pfnGetThemeMargins;
extern LPSETWINDOWTHEME pfnSetWindowTheme;
extern LPGETTHEMESYSSIZE pfnGetThemeSysSize;
extern LPMSIPROVIDEQUALIFIEDCOMPONENT pfnMsiProvideQualifiedComponent;
extern LPMSIGETFILEVERSION pfnMsiGetFileVersion;
extern LPSHGETPROPERTYSTOREFORWINDOW pfnSHGetPropertyStoreForWindow;

_Check_return_ HMODULE LoadFromSystemDir(_In_z_ LPCTSTR szDLLName);
_Check_return_ HMODULE LoadFromOLMAPIDir(_In_z_ LPCTSTR szDLLName);

_Check_return_ HMODULE MyLoadLibraryA(_In_z_ LPCSTR lpszLibFileName);
_Check_return_ HMODULE MyLoadLibraryW(_In_z_ LPCWSTR lpszLibFileName);
#ifdef UNICODE
#define MyLoadLibrary MyLoadLibraryW
#else
#define MyLoadLibrary MyLoadLibraryA
#endif

void ImportProcs();

wstring GetMAPIPath(wstring szClient);

// Exported from StubUtils.cpp
HMODULE GetMAPIHandle();
void UnLoadPrivateMAPI();
void ForceOutlookMAPI(bool fForce);
void ForceSystemMAPI(bool fForce);
void SetMAPIHandle(HMODULE hinstMAPI);
HMODULE GetPrivateMAPI();
bool GetComponentPath(LPCSTR szComponent, LPSTR szQualifier, LPSTR szDllPath, DWORD cchBufferSize, bool fInstall);

// Keep this in sync with g_pszOutlookQualifiedComponents
#define oqcOfficeBegin 0
#define oqcOffice15 oqcOfficeBegin + 0
#define oqcOffice14 oqcOfficeBegin + 1
#define oqcOffice12 oqcOfficeBegin + 2
#define oqcOffice11 oqcOfficeBegin + 3
#define oqcOffice11Debug oqcOfficeBegin + 4
#define oqcOfficeEnd oqcOffice11Debug

extern WCHAR g_pszOutlookQualifiedComponents[][MAX_PATH];

// Looks up Outlook's path given its qualified component guid
LPWSTR GetOutlookPath(_In_z_ LPCWSTR szCategory, _Out_opt_ bool* lpb64);

enum mapiSource;
class MAPIPathIterator
{
public:
	MAPIPathIterator(bool bBypassRestrictions);
	~MAPIPathIterator();
	LPWSTR GetNextMAPIPath();
	LPWSTR GetInstalledOutlookMAPI(int iOutlook) const;
	LPWSTR GetMAPISystemDir() const;

private:
	LPWSTR GetRegisteredMapiClient(LPCWSTR pwzProviderOverride, bool bDLL, bool bEx);
	LPWSTR GetMailClientFromMSIData(HKEY hkeyMapiClient) const;
	static LPWSTR GetMailClientFromDllPath(HKEY hkeyMapiClient, bool bEx);
	LPWSTR GetNextInstalledOutlookMAPI();

	mapiSource CurrentSource;
	HKEY m_hMailKey;
	HKEY m_hkeyMapiClient;
	LPWSTR m_rgchMailClient;
	LPCWSTR m_szRegisteredClient;
	bool m_bBypassRestrictions;

	int m_iCurrentOutlook;
};

_Check_return_ STDAPI HrCopyRestriction(
	_In_ LPSRestriction lpResSrc, // source restriction ptr
	_In_opt_ LPVOID lpObject, // ptr to existing MAPI buffer
	_In_ LPSRestriction* lppResDest // dest restriction buffer ptr
);

_Check_return_ HRESULT HrCopyRestrictionArray(
	_In_ LPSRestriction lpResSrc, // source restriction
	_In_ LPVOID lpObject, // ptr to existing MAPI buffer
	ULONG cRes, // # elements in array
	_In_count_(cRes) LPSRestriction lpResDest // destination restriction
);

_Check_return_ STDMETHODIMP MyOpenStreamOnFile(_In_ LPALLOCATEBUFFER lpAllocateBuffer,
	_In_ LPFREEBUFFER lpFreeBuffer,
	ULONG ulFlags,
	_In_z_ LPCWSTR lpszFileName,
	_In_opt_z_ LPCWSTR /*lpszPrefix*/,
	_Out_ LPSTREAM* lppStream);

void WINAPI MyHeapSetInformation(_In_opt_ HANDLE HeapHandle,
	_In_ HEAP_INFORMATION_CLASS HeapInformationClass,
	_In_opt_count_(HeapInformationLength) PVOID HeapInformation,
	_In_ SIZE_T HeapInformationLength);

_Check_return_ STDAPI_(SCODE) MyPropCopyMore(_In_ LPSPropValue lpSPropValueDest,
	_In_ LPSPropValue lpSPropValueSrc,
	_In_ ALLOCATEMORE * lpfAllocMore,
	_In_ LPVOID lpvObject);

HRESULT WINAPI MyMimeOleGetCodePageCharset(
	CODEPAGEID cpiCodePage,
	CHARSETTYPE ctCsetType,
	LPHCHARSET phCharset);

BOOL WINAPI MyGetModuleHandleExW(
	DWORD dwFlags,
	LPCWSTR lpModuleName,
	HMODULE* phModule);