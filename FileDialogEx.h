#pragma once

// Make sure OPENFILENAMEEX is the same size regardless of how _WIN32_WINNT is defined
#if (_WIN32_WINNT >= 0x0500)
struct OPENFILENAMEEXA : public OPENFILENAMEA {};
struct OPENFILENAMEEXW : public OPENFILENAMEW {};
#else
// Windows 2000 version of OPENFILENAME.
// The new version has three extra members.
// This is copied from commdlg.h
struct OPENFILENAMEEXA : public OPENFILENAMEA
{
	void*         pvReserved;
	DWORD         dwReserved;
	DWORD         FlagsEx;
};
struct OPENFILENAMEEXW : public OPENFILENAMEW
{
	void*         pvReserved;
	DWORD         dwReserved;
	DWORD         FlagsEx;
};
#endif

// CFileDialogExA: Encapsulate Windows-2000 style open dialog.
class CFileDialogExA{
public:
	CFileDialogExA();
	~CFileDialogExA();

	_Check_return_ INT_PTR DisplayDialog(bool bOpenFileDialog, // true for open, false for FileSaveAs
		_In_opt_z_ LPCSTR lpszDefExt = NULL,
		_In_opt_z_ LPCSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_In_opt_z_ LPCSTR lpszFilter = NULL,
		_In_opt_ CWnd* pParentWnd = NULL);

public:
	// The values returned here are only valid as long as this object is valid
	// To keep them, make a copy before this object is destroyed
	_Check_return_ LPSTR GetFileName();
	_Check_return_ LPSTR GetNextFileName();

private:
	bool m_bOpenFileDialog;
	CHAR m_szFileName[_MAX_PATH]; // contains full path name after return

	OPENFILENAMEEXA	m_ofn; // Windows 2000 version of OPENFILENAME
	LPSTR			m_szBigBuff;
	LPSTR			m_szNextPath;
	CHAR			strPath[_MAX_PATH];
};

// CFileDialogExA: Encapsulate Windows-2000 style open dialog.
class CFileDialogExW{
public:
	CFileDialogExW();
	~CFileDialogExW();

	_Check_return_ INT_PTR DisplayDialog(bool bOpenFileDialog, // true for open, false for FileSaveAs
		_In_opt_z_ LPCWSTR lpszDefExt = NULL,
		_In_opt_z_ LPCWSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_In_ wstring lpszFilter = wstring(L""),
		_In_opt_ CWnd* pParentWnd = NULL);

public:
	// The values returned here are only valid as long as this object is valid
	// To keep them, make a copy before this object is destroyed
	_Check_return_ LPWSTR GetFileName();
	_Check_return_ LPWSTR GetNextFileName();

private:
	bool m_bOpenFileDialog;
	WCHAR m_szFileName[_MAX_PATH]; // contains full path name after return

	OPENFILENAMEEXW	m_ofn; // Windows 2000 version of OPENFILENAME
	LPWSTR			m_szBigBuff;
	LPWSTR			m_szNextPath;
	WCHAR			strPath[_MAX_PATH];
};

#ifdef UNICODE
#define CFileDialogEx CFileDialogExW
#else
#define CFileDialogEx CFileDialogExA
#endif