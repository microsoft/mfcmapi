#pragma once

// Make sure OPENFILENAMEEX is the same size regardless of how _WIN32_WINNT is defined
#if (_WIN32_WINNT >= 0x0500)
struct OPENFILENAMEEXW : public OPENFILENAMEW {};
#else
struct OPENFILENAMEEXW : public OPENFILENAMEW
{
	void* pvReserved;
	DWORD dwReserved;
	DWORD FlagsEx;
};
#endif

// CFileDialogExW: Encapsulate Windows-2000 style open dialog.
class CFileDialogExW
{
public:
	CFileDialogExW();
	~CFileDialogExW();

	_Check_return_ INT_PTR DisplayDialog(bool bOpenFileDialog, // true for open, false for FileSaveAs
		_In_opt_z_ LPCWSTR lpszDefExt = nullptr,
		_In_opt_z_ LPCWSTR lpszFileName = nullptr,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_In_ wstring lpszFilter = wstring(L""),
		_In_opt_ CWnd* pParentWnd = nullptr);

public:
	// The values returned here are only valid as long as this object is valid
	// To keep them, make a copy before this object is destroyed
	_Check_return_ LPWSTR GetFileName() const;
	_Check_return_ LPWSTR GetNextFileName();

private:
	bool m_bOpenFileDialog;
	WCHAR m_szFileName[_MAX_PATH]; // contains full path name after return

	OPENFILENAMEEXW m_ofn; // Windows 2000 version of OPENFILENAME
	LPWSTR m_szBigBuff;
	LPWSTR m_szNextPath;
	WCHAR strPath[_MAX_PATH];
};