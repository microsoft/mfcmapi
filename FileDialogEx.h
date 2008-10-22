#pragma once
// FileDialogEx.h : Extended file dialog class to work around issues in the base MFC class.

// Originally from MSDN August 2000, by Paul DiLascia.
// With substantial changes to work across more versions of Windows

// Make sure OPENFILENAMEEX is the same size regardless of how _WIN32_WINNT is defined
#if (_WIN32_WINNT >= 0x0500)
struct OPENFILENAMEEX : public OPENFILENAME {};
#else
// Windows 2000 version of OPENFILENAME.
// The new version has three extra members.
// This is copied from commdlg.h
struct OPENFILENAMEEX : public OPENFILENAME {
	void*         pvReserved;
	DWORD         dwReserved;
	DWORD         FlagsEx;
};
#endif

// Copied from commdlg.h - size of the 'NT4 version' of OPENFILENAME
// Not always defined, but we need to know this size
#ifndef OPENFILENAME_SIZE_VERSION_400
#define OPENFILENAME_SIZE_VERSION_400A  CDSIZEOF_STRUCT(OPENFILENAMEA,lpTemplateName)
#define OPENFILENAME_SIZE_VERSION_400W  CDSIZEOF_STRUCT(OPENFILENAMEW,lpTemplateName)
#ifdef UNICODE
#define OPENFILENAME_SIZE_VERSION_400  OPENFILENAME_SIZE_VERSION_400W
#else
#define OPENFILENAME_SIZE_VERSION_400  OPENFILENAME_SIZE_VERSION_400A
#endif // !UNICODE
#endif // OPENFILENAME_SIZE_VERSION_400

// CFileDialogEx: Encapsulate Windows-2000 style open dialog.
class CFileDialogEx : public CFileDialog {
public:
	CFileDialogEx(BOOL bOpenFileDialog, // TRUE for open, FALSE for FileSaveAs
		LPCTSTR lpszDefExt = NULL,
		LPCTSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		LPCTSTR lpszFilter = NULL,
		CWnd* pParentWnd = NULL);
	virtual ~CFileDialogEx();

	INT_PTR DoModal();

private:
	BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult);

	OPENFILENAMEEX	m_ofnEx; // Windows 2000 version of OPENFILENAME
	LPTSTR			m_szbigBuff;
};
