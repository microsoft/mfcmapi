// Loading imports from DLLs
#pragma once

// For EditSecurity
typedef bool(STDAPICALLTYPE EDITSECURITY)(HWND hwndOwner, LPSECURITYINFO psi);
typedef EDITSECURITY* LPEDITSECURITY;

// For StgCreateStorageEx
typedef HRESULT(STDAPICALLTYPE STGCREATESTORAGEEX)(
	IN const WCHAR* pwcsName,
	IN DWORD grfMode,
	IN DWORD stgfmt, // enum
	IN DWORD grfAttrs, // reserved
	IN STGOPTIONS* pStgOptions,
	IN void* reserved,
	IN REFIID riid,
	OUT void** ppObjectOpen);
typedef STGCREATESTORAGEEX* LPSTGCREATESTORAGEEX;

// For Themes
typedef HTHEME(STDMETHODCALLTYPE OPENTHEMEDATA)(HWND hwnd, LPCWSTR pszClassList);
typedef OPENTHEMEDATA* LPOPENTHEMEDATA;

typedef HTHEME(STDMETHODCALLTYPE CLOSETHEMEDATA)(HTHEME hTheme);
typedef CLOSETHEMEDATA* LPCLOSETHEMEDATA;

typedef HRESULT(STDMETHODCALLTYPE GETTHEMEMARGINS)(
	HTHEME hTheme,
	OPTIONAL HDC hdc,
	int iPartId,
	int iStateId,
	int iPropId,
	OPTIONAL RECT* prc,
	OUT MARGINS* pMargins);
typedef GETTHEMEMARGINS* LPGETTHEMEMARGINS;

typedef HRESULT(
	STDMETHODCALLTYPE SETWINDOWTHEME)(__in HWND hwnd, __in LPCWSTR pszSubAppName, __in LPCWSTR pszSubIdList);
typedef SETWINDOWTHEME* LPSETWINDOWTHEME;

typedef int(STDMETHODCALLTYPE GETTHEMESYSSIZE)(HTHEME hTheme, int iSizeID);
typedef GETTHEMESYSSIZE* LPGETTHEMESYSSIZE;

typedef HRESULT(STDMETHODCALLTYPE MSIGETFILEVERSION)(
	LPCWSTR szFilePath,
	LPWSTR lpVersionBuf,
	LPDWORD pcchVersionBuf,
	LPWSTR lpLangBuf,
	LPDWORD pcchLangBuf);
typedef MSIGETFILEVERSION* LPMSIGETFILEVERSION;

typedef HRESULT(STDMETHODCALLTYPE SHGETPROPERTYSTOREFORWINDOW)(HWND hwnd, REFIID riid, void** ppv);
typedef SHGETPROPERTYSTOREFORWINDOW* LPSHGETPROPERTYSTOREFORWINDOW;

typedef LONG(STDMETHODCALLTYPE FINDPACKAGESBYPACKAGEFAMILY)(
	PCWSTR packageFamilyName,
	UINT32 packageFilters,
	UINT32* count,
	PWSTR* packageFullNames,
	UINT32* bufferLength,
	WCHAR* buffer,
	UINT32* packageProperties);
typedef FINDPACKAGESBYPACKAGEFAMILY* LPFINDPACKAGESBYPACKAGEFAMILY;

typedef LONG(STDMETHODCALLTYPE
				 PACKAGEIDFROMFULLNAME)(PCWSTR packageFullName, const UINT32 flags, UINT32* bufferLength, BYTE* buffer);
typedef PACKAGEIDFROMFULLNAME* LPPACKAGEIDFROMFULLNAME;

typedef DPAPI_IMP BOOL(WINAPI CRYPTPROTECTDATA)(
	DATA_BLOB* pDataIn,
	LPCWSTR szDataDescr,
	DATA_BLOB* pOptionalEntropy,
	PVOID pvReserved,
	CRYPTPROTECT_PROMPTSTRUCT* pPromptStruct,
	DWORD dwFlags,
	DATA_BLOB* pDataOut
);
typedef CRYPTPROTECTDATA* LPCRYPTPROTECTDATA;
typedef BOOL(WINAPI CRYPTUNPROTECTDATA)(
	DATA_BLOB* pDataIn,
	LPWSTR* ppszDataDescr,
	DATA_BLOB* pOptionalEntropy,
	PVOID pvReserved,
	CRYPTPROTECT_PROMPTSTRUCT* pPromptStruct,
	DWORD dwFlags,
	DATA_BLOB* pDataOut);
typedef CRYPTUNPROTECTDATA* LPCRYPTUNPROTECTDATA;

namespace import
{
	extern LPEDITSECURITY pfnEditSecurity;
	extern LPSTGCREATESTORAGEEX pfnStgCreateStorageEx;
	extern LPOPENTHEMEDATA pfnOpenThemeData;
	extern LPCLOSETHEMEDATA pfnCloseThemeData;
	extern LPGETTHEMEMARGINS pfnGetThemeMargins;
	extern LPSETWINDOWTHEME pfnSetWindowTheme;
	extern LPGETTHEMESYSSIZE pfnGetThemeSysSize;
	extern LPMSIGETFILEVERSION pfnMsiGetFileVersion;
	extern LPSHGETPROPERTYSTOREFORWINDOW pfnSHGetPropertyStoreForWindow;
	extern LPFINDPACKAGESBYPACKAGEFAMILY pfnFindPackagesByPackageFamily;
	extern LPPACKAGEIDFROMFULLNAME pfnPackageIdFromFullName;
	extern LPCRYPTPROTECTDATA pfnCryptProtectData;
	extern LPCRYPTUNPROTECTDATA pfnCryptUnprotectData;

	_Check_return_ HMODULE LoadFromSystemDir(_In_ const std::wstring& szDLLName);
	_Check_return_ HMODULE MyLoadLibraryW(_In_ const std::wstring& lpszLibFileName);

	void ImportProcs();

	std::wstring GetMAPIPath(const std::wstring& szClient);

	void WINAPI MyHeapSetInformation(
		_In_opt_ HANDLE HeapHandle,
		_In_ HEAP_INFORMATION_CLASS HeapInformationClass,
		_In_opt_count_(HeapInformationLength) PVOID HeapInformation,
		_In_ SIZE_T HeapInformationLength);

	HRESULT WINAPI MyMimeOleGetCodePageCharset(CODEPAGEID cpiCodePage, CHARSETTYPE ctCsetType, LPHCHARSET phCharset);
} // namespace import