#include "stdafx.h"
#include "FileDialogEx.h"

_Check_return_ static BOOL IsWin2000();

///////////////////////////////////////////////////////////////////////////
// CFileDialogExA / CFileDialogExW

#define CCHBIGBUFF 8192

CFileDialogExA::CFileDialogExA()
{
	m_szBigBuff = NULL;
	m_szNextPath = NULL;
} // CFileDialogExA::CFileDialogExA

CFileDialogExA::~CFileDialogExA()
{
	delete[] m_szBigBuff;
} // CFileDialogExA::~CFileDialogExA

_Check_return_ INT_PTR CFileDialogExA::DisplayDialog(BOOL bOpenFileDialog, // TRUE for open, FALSE for FileSaveAs
									  _In_opt_z_ LPCSTR lpszDefExt,
									  _In_opt_z_ LPCSTR lpszFileName,
									  DWORD dwFlags,
									  _In_opt_z_ LPCSTR lpszFilter,
									  _In_opt_ CWnd* pParentWnd)
{
	m_szBigBuff = NULL;
	m_szNextPath = NULL;
	m_bOpenFileDialog = bOpenFileDialog;

	// initialize structure to 0/NULL
	memset(&m_ofn, 0, sizeof(OPENFILENAMEEXA));

	if (IsWin2000())
		m_ofn.lStructSize = sizeof(OPENFILENAMEEXA);
	else
		m_ofn.lStructSize = sizeof(OPENFILENAMEA);

	m_ofn.lpstrFile = m_szFileName;
	m_ofn.nMaxFile = _countof(m_szFileName);

	if (dwFlags & OFN_ALLOWMULTISELECT)
	{
		if (m_szBigBuff) delete[] m_szBigBuff;
		m_szBigBuff = new CHAR[CCHBIGBUFF];
		if (m_szBigBuff)
		{
			m_szBigBuff[0] = 0; // NULL terminate
			m_ofn.lpstrFile = m_szBigBuff;
			m_ofn.nMaxFile = CCHBIGBUFF;
		}
	}

	m_ofn.lpstrDefExt = lpszDefExt;
	m_ofn.Flags = dwFlags | OFN_ENABLEHOOK | OFN_EXPLORER;
	m_ofn.hwndOwner = pParentWnd?pParentWnd->m_hWnd:NULL;

	// zero out the file buffer for consistent parsing later
	memset(m_ofn.lpstrFile, 0, m_ofn.nMaxFile);

	// setup initial file name
	if (lpszFileName != NULL)
		strncpy_s(m_szFileName, _countof(m_szFileName), lpszFileName, _TRUNCATE);

	// Translate filter into commdlg format (lots of \0)
	CStringA strFilter; // filter string
	if (lpszFilter != NULL)
	{
		strFilter = lpszFilter;
		LPSTR pch = strFilter.GetBuffer(0); // modify the buffer in place
		// MFC delimits with '|' not '\0'
		while ((pch = strchr(pch, '|')) != NULL)
			*pch++ = '\0';
		m_ofn.lpstrFilter = strFilter;
	}

	int nResult;
	if (m_bOpenFileDialog)
		nResult = ::GetOpenFileNameA((OPENFILENAMEA*)&m_ofn);
	else
		nResult = ::GetSaveFileNameA((OPENFILENAMEA*)&m_ofn);

	m_szNextPath = m_ofn.lpstrFile;
	return nResult ? nResult : IDCANCEL;
} // CFileDialogExA::DisplayDialog

_Check_return_ LPSTR CFileDialogExA::GetFileName()
{
	return m_ofn.lpstrFile;
} // CFileDialogExA::GetFileName

_Check_return_ LPSTR CFileDialogExA::GetNextFileName()
{
	if (!m_szNextPath) return NULL;
	LPSTR lpsz = m_szNextPath;
	if (lpsz == m_ofn.lpstrFile) // first time
	{
		if ((m_ofn.Flags & OFN_ALLOWMULTISELECT) == 0)
		{
			m_szNextPath = NULL;
			return m_ofn.lpstrFile;
		}

		// find char pos after first Delimiter
		while(*lpsz != '\0') lpsz = CharNextA(lpsz);
		lpsz = CharNextA(lpsz);

		// if single selection then return only selection
		if (*lpsz == 0)
		{
			m_szNextPath = NULL;
			return m_ofn.lpstrFile;
		}
	}

	CStringA strBasePath = m_ofn.lpstrFile;
	LPSTR strFileName = lpsz;

	// find char pos at next Delimiter
	while(*lpsz != '\0') lpsz = CharNextA(lpsz);
	lpsz = CharNextA(lpsz);

	if (*lpsz == '\0') // if double terminated then done
		m_szNextPath = NULL;
	else
		m_szNextPath = lpsz;

	CHAR strDrive[_MAX_DRIVE], strDir[_MAX_DIR], strName[_MAX_FNAME], strExt[_MAX_EXT];
	_splitpath_s(strFileName, strDrive, _MAX_DRIVE, strDir, _MAX_DIR, strName, _MAX_FNAME, strExt, _MAX_EXT);
	if (*strDrive || *strDir)
	{
		strcpy_s(strPath, _countof(strPath), strFileName);
	}
	else
	{
		_splitpath_s(strBasePath+"\\", strDrive, _MAX_DRIVE, strDir, _MAX_DIR, NULL, 0, NULL, 0); // STRING_OK
		_makepath_s(strPath, _MAX_PATH, strDrive, strDir, strName, strExt);
	}

	return strPath;
} // CFileDialogExA::GetNextFileName

CFileDialogExW::CFileDialogExW()
{
	m_szBigBuff = NULL;
	m_szNextPath = NULL;
} // CFileDialogExW::CFileDialogExW

CFileDialogExW::~CFileDialogExW()
{
	delete[] m_szBigBuff;
} // CFileDialogExW::~CFileDialogExW

_Check_return_ INT_PTR CFileDialogExW::DisplayDialog(BOOL bOpenFileDialog, // TRUE for open, FALSE for FileSaveAs
									  _In_opt_z_ LPCWSTR lpszDefExt,
									  _In_opt_z_ LPCWSTR lpszFileName,
									  DWORD dwFlags,
									  _In_opt_z_ LPCWSTR lpszFilter,
									  _In_opt_ CWnd* pParentWnd)
{
	m_szBigBuff = NULL;
	m_szNextPath = NULL;
	m_bOpenFileDialog = bOpenFileDialog;

	// initialize structure to 0/NULL
	memset(&m_ofn, 0, sizeof(OPENFILENAMEEXW));

	if (IsWin2000())
		m_ofn.lStructSize = sizeof(OPENFILENAMEEXW);
	else
		m_ofn.lStructSize = sizeof(OPENFILENAMEW);

	m_ofn.lpstrFile = m_szFileName;
	m_ofn.nMaxFile = _countof(m_szFileName);

	if (dwFlags & OFN_ALLOWMULTISELECT)
	{
		if (m_szBigBuff) delete[] m_szBigBuff;
		m_szBigBuff = new WCHAR[CCHBIGBUFF];
		if (m_szBigBuff)
		{
			m_szBigBuff[0] = 0; // NULL terminate
			m_ofn.lpstrFile = m_szBigBuff;
			m_ofn.nMaxFile = CCHBIGBUFF;
		}
	}

	m_ofn.lpstrDefExt = lpszDefExt;
	m_ofn.Flags = dwFlags | OFN_ENABLEHOOK | OFN_EXPLORER;
	m_ofn.hwndOwner = pParentWnd?pParentWnd->m_hWnd:NULL;

	// zero out the file buffer for consistent parsing later
	memset(m_ofn.lpstrFile, 0, m_ofn.nMaxFile);

	// setup initial file name
	if (lpszFileName != NULL)
		wcsncpy_s(m_szFileName, _countof(m_szFileName), lpszFileName, _TRUNCATE);

	// Translate filter into commdlg format (lots of \0)
	CStringW strFilter; // filter string
	if (lpszFilter != NULL)
	{
		strFilter = lpszFilter;
		LPWSTR pch = strFilter.GetBuffer(0); // modify the buffer in place
		// MFC delimits with '|' not '\0'
		while ((pch = wcschr(pch, L'|')) != NULL)
			*pch++ = L'\0';
		m_ofn.lpstrFilter = strFilter;
	}

	int nResult;
	if (m_bOpenFileDialog)
		nResult = ::GetOpenFileNameW((OPENFILENAMEW*)&m_ofn);
	else
		nResult = ::GetSaveFileNameW((OPENFILENAMEW*)&m_ofn);

	m_szNextPath = m_ofn.lpstrFile;
	return nResult ? nResult : IDCANCEL;
} // CFileDialogExW::DisplayDialog

_Check_return_ LPWSTR CFileDialogExW::GetFileName()
{
	return m_ofn.lpstrFile;
} // CFileDialogExW::GetFileName

_Check_return_ LPWSTR CFileDialogExW::GetNextFileName()
{
	if (!m_szNextPath) return NULL;
	LPWSTR lpsz = m_szNextPath;
	if (lpsz == m_ofn.lpstrFile) // first time
	{
		if ((m_ofn.Flags & OFN_ALLOWMULTISELECT) == 0)
		{
			m_szNextPath = NULL;
			return m_ofn.lpstrFile;
		}

		// find char pos after first Delimiter
		while(*lpsz != L'\0') lpsz = CharNextW(lpsz);
		lpsz = CharNextW(lpsz);

		// if single selection then return only selection
		if (*lpsz == 0)
		{
			m_szNextPath = NULL;
			return m_ofn.lpstrFile;
		}
	}

	CStringW strBasePath = m_ofn.lpstrFile;
	LPWSTR strFileName = lpsz;

	// find char pos at next Delimiter
	while(*lpsz != L'\0') lpsz = CharNextW(lpsz);
	lpsz = CharNextW(lpsz);

	if (*lpsz == L'\0') // if double terminated then done
		m_szNextPath = NULL;
	else
		m_szNextPath = lpsz;

	WCHAR strDrive[_MAX_DRIVE], strDir[_MAX_DIR], strName[_MAX_FNAME], strExt[_MAX_EXT];
	_wsplitpath_s(strFileName, strDrive, _MAX_DRIVE, strDir, _MAX_DIR, strName, _MAX_FNAME, strExt, _MAX_EXT);
	if (*strDrive || *strDir)
	{
		wcscpy_s(strPath, _countof(strPath), strFileName);
	}
	else
	{
		_wsplitpath_s(strBasePath+L"\\", strDrive, _MAX_DRIVE, strDir, _MAX_DIR, NULL, 0, NULL, 0); // STRING_OK
		_wmakepath_s(strPath, _MAX_PATH, strDrive, strDir, strName, strExt);
	}

	return strPath;
} // CFileDialogExW::GetNextFileName

_Check_return_ BOOL IsWin2000()
{
	OSVERSIONINFOEX osvi;
	HRESULT hRes = S_OK;

	// Try calling GetVersionEx using the OSVERSIONINFOEX structure,
	// which is supported on Windows 2000.
	//
	// If that fails, try using the OSVERSIONINFO structure.

	ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);

	WC_B(GetVersionEx ((OSVERSIONINFO*) &osvi));
	if (S_OK != hRes)
	{
		// If OSVERSIONINFOEX doesn't work, try OSVERSIONINFO.
		hRes = S_OK;
		osvi.dwOSVersionInfoSize = sizeof (OSVERSIONINFO);
		EC_B(GetVersionEx((OSVERSIONINFO*) &osvi));
		if (S_OK != hRes) return FALSE;
	}

	switch (osvi.dwPlatformId)
	{
	case VER_PLATFORM_WIN32_NT:

		if (osvi.dwMajorVersion >= 5)
			return TRUE;

		break;
	}
	return FALSE;
} // IsWin2000