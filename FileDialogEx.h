#pragma once

// Make sure OPENFILENAMEEX is the same size regardless of how _WIN32_WINNT is defined
#if (_WIN32_WINNT >= 0x0500)
struct OPENFILENAMEEXA : public OPENFILENAMEA {};
struct OPENFILENAMEEXW : public OPENFILENAMEW {};
#else
// Windows 2000 version of OPENFILENAME.
// The new version has three extra members.
// This is copied from commdlg.h
struct OPENFILENAMEEXA : public OPENFILENAMEA {
	void*         pvReserved;
	DWORD         dwReserved;
	DWORD         FlagsEx;
};
struct OPENFILENAMEEXW : public OPENFILENAMEW {
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

	INT_PTR DisplayDialog(BOOL bOpenFileDialog, // TRUE for open, FALSE for FileSaveAs
		LPCSTR lpszDefExt = NULL,
		LPCSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		LPCSTR lpszFilter = NULL,
		CWnd* pParentWnd = NULL);

public:
	// The values returned here are only valid as long as this object is valid
	// To keep them, make a copy before this object is destroyed
	LPSTR GetFileName();
	LPSTR GetNextFileName();

private:
	BOOL m_bOpenFileDialog;
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

	INT_PTR DisplayDialog(BOOL bOpenFileDialog, // TRUE for open, FALSE for FileSaveAs
		LPCWSTR lpszDefExt = NULL,
		LPCWSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		LPCWSTR lpszFilter = NULL,
		CWnd* pParentWnd = NULL);

public:
	// The values returned here are only valid as long as this object is valid
	// To keep them, make a copy before this object is destroyed
	LPWSTR GetFileName();
	LPWSTR GetNextFileName();

private:
	BOOL m_bOpenFileDialog;
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