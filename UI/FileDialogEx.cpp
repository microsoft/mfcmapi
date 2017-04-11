#include "stdafx.h"
#include <UI/FileDialogEx.h>

wstring CFileDialogExW::OpenFile(
	_In_ const wstring& lpszDefExt,
	_In_ const wstring& lpszFileName,
	DWORD dwFlags,
	_In_ const wstring& lpszFilter,
	_In_opt_ CWnd* pParentWnd)
{
	auto hRes = S_OK;
	INT_PTR iDlgRet = IDOK;
	CFileDialogExW dlgFilePicker;
	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		true,
		lpszDefExt,
		lpszFileName,
		dwFlags,
		lpszFilter,
		pParentWnd));

	if (iDlgRet == IDOK)
	{
		return dlgFilePicker.GetFileName();
	}

	return wstring();
}

vector<wstring> CFileDialogExW::OpenFiles(
	_In_ const wstring& lpszDefExt,
	_In_ const wstring& lpszFileName,
	DWORD dwFlags,
	_In_ const wstring& lpszFilter,
	_In_opt_ CWnd* pParentWnd)
{
	auto hRes = S_OK;
	INT_PTR iDlgRet = IDOK;
	CFileDialogExW dlgFilePicker;
	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		true,
		lpszDefExt,
		lpszFileName,
		dwFlags | OFN_ALLOWMULTISELECT,
		lpszFilter,
		pParentWnd));

	if (iDlgRet == IDOK)
	{
		return dlgFilePicker.GetFileNames();
	}

	return vector<wstring>();
}

wstring CFileDialogExW::SaveAs(
	_In_ const wstring& lpszDefExt,
	_In_ const wstring& lpszFileName,
	DWORD dwFlags,
	_In_ const wstring& lpszFilter,
	_In_opt_ CWnd* pParentWnd)
{
	auto hRes = S_OK;
	INT_PTR iDlgRet = IDOK;
	CFileDialogExW dlgFilePicker;
	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		false,
		lpszDefExt,
		lpszFileName,
		dwFlags,
		lpszFilter,
		pParentWnd));

	if (iDlgRet == IDOK)
	{
		return dlgFilePicker.GetFileName();
}

	return wstring();
}

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

#define CCHBIGBUFF 8192

vector<wstring> UnpackFileNames(OPENFILENAMEEXW ofn)
{
	auto paths = vector<wstring>();

	if ((ofn.Flags & OFN_ALLOWMULTISELECT) == 0)
	{
		paths.push_back(ofn.lpstrFile);
		return paths;
	}

	auto strBasePath = wstring(ofn.lpstrFile) + L"\\";
	auto lpsz = ofn.lpstrFile;
	// find char pos after first Delimiter
	while (*lpsz != L'\0') lpsz = CharNextW(lpsz);
	lpsz++;

	// if single selection then return only selection
	if (*lpsz == L'\0')
	{
		paths.push_back(ofn.lpstrFile);
		return paths;
	}

	while (lpsz != nullptr)
	{
		WCHAR strPath[_MAX_PATH], strDrive[_MAX_DRIVE], strDir[_MAX_DIR], strName[_MAX_FNAME], strExt[_MAX_EXT];
		_wsplitpath_s(lpsz, strDrive, _MAX_DRIVE, strDir, _MAX_DIR, strName, _MAX_FNAME, strExt, _MAX_EXT);
		if (*strDrive || *strDir)
		{
			wcscpy_s(strPath, _countof(strPath), lpsz);
		}
		else
		{
			_wsplitpath_s(strBasePath.c_str(), strDrive, _MAX_DRIVE, strDir, _MAX_DIR, nullptr, 0, nullptr, 0); // STRING_OK
			_wmakepath_s(strPath, _MAX_PATH, strDrive, strDir, strName, strExt);
		}

		paths.push_back(strPath);

		// find char pos at next Delimiter
		while (*lpsz != L'\0') lpsz = CharNextW(lpsz);
		lpsz++;

		if (*lpsz == L'\0') // if double terminated then done
			lpsz = nullptr;
	}

	return paths;
}

_Check_return_ INT_PTR CFileDialogExW::DisplayDialog(
	bool bOpenFileDialog, // true for open, false for FileSaveAs
	_In_ const wstring& lpszDefExt,
	_In_ const wstring& lpszFileName,
	DWORD dwFlags,
	_In_ const wstring& lpszFilter,
	_In_opt_ CWnd* pParentWnd)
{
	WCHAR szFileName[_MAX_PATH]; // contains full path name after return

	// initialize structure
	OPENFILENAMEEXW ofn;
	memset(&ofn, 0, sizeof(OPENFILENAMEEXW));
	ofn.lStructSize = sizeof(OPENFILENAMEEXW);
	ofn.lpstrFile = szFileName;
	ofn.nMaxFile = _countof(szFileName);

	LPWSTR szBigBuff = nullptr;
	if (dwFlags & OFN_ALLOWMULTISELECT)
	{
		szBigBuff = new WCHAR[CCHBIGBUFF];
		if (szBigBuff)
		{
			szBigBuff[0] = 0; // NULL terminate
			ofn.lpstrFile = szBigBuff;
			ofn.nMaxFile = CCHBIGBUFF;
		}
	}

	ofn.lpstrDefExt = lpszDefExt.c_str();
	ofn.Flags = dwFlags | OFN_ENABLEHOOK | OFN_EXPLORER;
	ofn.hwndOwner = pParentWnd ? pParentWnd->m_hWnd : NULL;

	// zero out the file buffer for consistent parsing later
	memset(ofn.lpstrFile, 0, ofn.nMaxFile);

	// setup initial file name
	if (!lpszFileName.empty())
		wcsncpy_s(ofn.lpstrFile, ofn.nMaxFile, lpszFileName.c_str(), _TRUNCATE);

	// Translate filter into commdlg format (lots of \0)
	auto strFilter = lpszFilter; // filter string
	if (!strFilter.empty())
	{
		replace(strFilter.begin(), strFilter.end(), L'|', L'\0');
		ofn.lpstrFilter = strFilter.c_str();
	}

	BOOL bResult;
	if (bOpenFileDialog)
		bResult = GetOpenFileNameW(static_cast<OPENFILENAMEW*>(&ofn));
	else
		bResult = GetSaveFileNameW(static_cast<OPENFILENAMEW*>(&ofn));

	if (bResult)
	{
		m_paths = UnpackFileNames(ofn);
	}
	
	delete[] szBigBuff;
	return bResult ? bResult : IDCANCEL;
}

wstring CFileDialogExW::GetFileName() const
{
	if (m_paths.size() >= 1) return m_paths[0];
	return emptystring;
}

vector<wstring> CFileDialogExW::GetFileNames() const
{
	return m_paths;
}