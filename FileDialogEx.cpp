#include "stdafx.h"
#include "FileDialogEx.h"
#include <algorithm>

#define CCHBIGBUFF 8192

CFileDialogExW::CFileDialogExW(): m_bOpenFileDialog(false) {
	m_szBigBuff = nullptr;
	m_szNextPath = nullptr;
}

CFileDialogExW::~CFileDialogExW()
{
	delete[] m_szBigBuff;
}

_Check_return_ INT_PTR CFileDialogExW::DisplayDialog(bool bOpenFileDialog, // true for open, false for FileSaveAs
	_In_opt_z_ LPCWSTR lpszDefExt,
	_In_opt_z_ LPCWSTR lpszFileName,
	DWORD dwFlags,
	_In_ wstring lpszFilter,
	_In_opt_ CWnd* pParentWnd)
{
	m_szBigBuff = nullptr;
	m_szNextPath = nullptr;
	m_bOpenFileDialog = bOpenFileDialog;

	// initialize structure to 0/NULL
	memset(&m_ofn, 0, sizeof(OPENFILENAMEEXW));
	m_ofn.lStructSize = sizeof(OPENFILENAMEEXW);
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
	m_ofn.hwndOwner = pParentWnd ? pParentWnd->m_hWnd : NULL;

	// zero out the file buffer for consistent parsing later
	memset(m_ofn.lpstrFile, 0, m_ofn.nMaxFile);

	// setup initial file name
	if (lpszFileName != nullptr)
		wcsncpy_s(m_szFileName, _countof(m_szFileName), lpszFileName, _TRUNCATE);

	// Translate filter into commdlg format (lots of \0)
	wstring strFilter; // filter string
	if (!lpszFilter.empty())
	{
		std::replace(lpszFilter.begin(), lpszFilter.end(), L'|', L'\0');
		m_ofn.lpstrFilter = lpszFilter.c_str();
	}

	int nResult;
	if (m_bOpenFileDialog)
		nResult = ::GetOpenFileNameW(static_cast<OPENFILENAMEW*>(&m_ofn));
	else
		nResult = ::GetSaveFileNameW(static_cast<OPENFILENAMEW*>(&m_ofn));

	m_szNextPath = m_ofn.lpstrFile;
	return nResult ? nResult : IDCANCEL;
}

_Check_return_ LPWSTR CFileDialogExW::GetFileName() const
{
	return m_ofn.lpstrFile;
}

_Check_return_ LPWSTR CFileDialogExW::GetNextFileName()
{
	if (!m_szNextPath) return nullptr;
	auto lpsz = m_szNextPath;
	if (lpsz == m_ofn.lpstrFile) // first time
	{
		if ((m_ofn.Flags & OFN_ALLOWMULTISELECT) == 0)
		{
			m_szNextPath = nullptr;
			return m_ofn.lpstrFile;
		}

		// find char pos after first Delimiter
		while (*lpsz != L'\0') lpsz = CharNextW(lpsz);
		lpsz++;

		// if single selection then return only selection
		if (*lpsz == L'\0')
		{
			m_szNextPath = nullptr;
			return m_ofn.lpstrFile;
		}
	}

	CStringW strBasePath = m_ofn.lpstrFile;
	auto strFileName = lpsz;

	// find char pos at next Delimiter
	while (*lpsz != L'\0') lpsz = CharNextW(lpsz);
	lpsz++;

	if (*lpsz == L'\0') // if double terminated then done
		m_szNextPath = nullptr;
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
		_wsplitpath_s(strBasePath + L"\\", strDrive, _MAX_DRIVE, strDir, _MAX_DIR, nullptr, 0, nullptr, 0); // STRING_OK
		_wmakepath_s(strPath, _MAX_PATH, strDrive, strDir, strName, strExt);
	}

	return strPath;
}