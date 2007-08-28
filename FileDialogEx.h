// MSDN August 2000
// If this code works, it was written by Paul DiLascia. If not, I don't
// know who wrote it. Compiles with Visual C++ 6.0, runs on Windows 98
// and probably Windows NT too.
#pragma once

// Make sure OPENFILENAMEEX is the same size regardless of how _WIN32_WINNT is defined
#if (_WIN32_WINNT >= 0x0500)
struct OPENFILENAMEEX : public OPENFILENAME {};
#else
// Windows 2000 version of OPENFILENAME.
// The new version has three extra members.
// This is copied from commdlg.h
struct OPENFILENAMEEX : public OPENFILENAME {
	void *        pvReserved;
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

///////////////////////////////////////////////////////////////////////////
// CFileDialogEx: Encapsulate Windows-2000 style open dialog.
class CFileDialogEx : public CFileDialog {
	DECLARE_DYNAMIC(CFileDialogEx)
public:
	CFileDialogEx(BOOL bOpenFileDialog, // TRUE for open,
		// FALSE for FileSaveAs
		LPCTSTR lpszDefExt = NULL,
		LPCTSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		LPCTSTR lpszFilter = NULL,
		CWnd* pParentWnd = NULL);
	~CFileDialogEx();

	// override
	virtual INT_PTR DoModal();

protected:
	OPENFILENAMEEX m_ofnEx; // new Windows 2000 version of OPENFILENAME

	virtual BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult);
	LPTSTR m_szbigBuff;

	DECLARE_MESSAGE_MAP()
		//{{AFX_MSG(CFileDialogEx)
		//}}AFX_MSG
};
