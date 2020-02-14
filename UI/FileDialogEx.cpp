#include <StdAfx.h>
#include <UI/FileDialogEx.h>
#include <core/utility/strings.h>

namespace file
{
	std::wstring CFileDialogExW::OpenFile(
		_In_ const std::wstring& lpszDefExt,
		_In_ const std::wstring& lpszFileName,
		DWORD dwFlags,
		_In_ const std::wstring& lpszFilter,
		_In_opt_ CWnd* pParentWnd)
	{
		CFileDialogExW dlgFilePicker;
		const auto iDlgRet =
			EC_D_DIALOG(dlgFilePicker.DisplayDialog(true, lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd));
		if (iDlgRet == IDOK)
		{
			return dlgFilePicker.GetFileName();
		}

		return std::wstring();
	}

	std::vector<std::wstring> CFileDialogExW::OpenFiles(
		_In_ const std::wstring& lpszDefExt,
		_In_ const std::wstring& lpszFileName,
		DWORD dwFlags,
		_In_ const std::wstring& lpszFilter,
		_In_opt_ CWnd* pParentWnd)
	{
		CFileDialogExW dlgFilePicker;
		const auto iDlgRet = EC_D_DIALOG(dlgFilePicker.DisplayDialog(
			true, lpszDefExt, lpszFileName, dwFlags | OFN_ALLOWMULTISELECT, lpszFilter, pParentWnd));
		if (iDlgRet == IDOK)
		{
			return dlgFilePicker.GetFileNames();
		}

		return std::vector<std::wstring>();
	}

	std::wstring CFileDialogExW::SaveAs(
		_In_ const std::wstring& lpszDefExt,
		_In_ const std::wstring& lpszFileName,
		DWORD dwFlags,
		_In_ const std::wstring& lpszFilter,
		_In_opt_ CWnd* pParentWnd)
	{
		CFileDialogExW dlgFilePicker;
		const auto iDlgRet =
			EC_D_DIALOG(dlgFilePicker.DisplayDialog(false, lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd));
		if (iDlgRet == IDOK)
		{
			return dlgFilePicker.GetFileName();
		}

		return std::wstring();
	}

	// Make sure OPENFILENAMEEX is the same size regardless of how _WIN32_WINNT is defined
#if (_WIN32_WINNT >= 0x0500)
	struct OPENFILENAMEEXW : public OPENFILENAMEW
	{
	};
#else
	struct OPENFILENAMEEXW : public OPENFILENAMEW
	{
		void* pvReserved;
		DWORD dwReserved;
		DWORD FlagsEx;
	};
#endif

#define CCHBIGBUFF 8192

	std::vector<std::wstring> UnpackFileNames(OPENFILENAMEEXW ofn)
	{
		auto paths = std::vector<std::wstring>();

		if ((ofn.Flags & OFN_ALLOWMULTISELECT) == 0)
		{
			paths.push_back(ofn.lpstrFile);
			return paths;
		}

		auto strBasePath = std::wstring(ofn.lpstrFile) + L"\\";
		auto lpsz = ofn.lpstrFile;
		// find char pos after first Delimiter
		while (*lpsz != L'\0')
			lpsz = CharNextW(lpsz);
		lpsz++;

		// if single selection then return only selection
		if (*lpsz == L'\0')
		{
			paths.push_back(ofn.lpstrFile);
			return paths;
		}

		while (lpsz != nullptr)
		{
			WCHAR strPath[_MAX_PATH] = {};
			WCHAR strDrive[_MAX_DRIVE] = {};
			WCHAR strDir[_MAX_DIR] = {};
			WCHAR strName[_MAX_FNAME] = {};
			WCHAR strExt[_MAX_EXT] = {};
			_wsplitpath_s(lpsz, strDrive, _MAX_DRIVE, strDir, _MAX_DIR, strName, _MAX_FNAME, strExt, _MAX_EXT);
			if (*strDrive || *strDir)
			{
				wcscpy_s(strPath, _countof(strPath), lpsz);
			}
			else
			{
				_wsplitpath_s(
					strBasePath.c_str(), strDrive, _MAX_DRIVE, strDir, _MAX_DIR, nullptr, 0, nullptr, 0); // STRING_OK
				_wmakepath_s(strPath, _MAX_PATH, strDrive, strDir, strName, strExt);
			}

			paths.push_back(strPath);

			// find char pos at next Delimiter
			while (*lpsz != L'\0')
				lpsz = CharNextW(lpsz);
			lpsz++;

			if (*lpsz == L'\0') // if double terminated then done
				lpsz = nullptr;
		}

		return paths;
	}

	_Check_return_ INT_PTR CFileDialogExW::DisplayDialog(
		bool bOpenFileDialog, // true for open, false for FileSaveAs
		_In_ const std::wstring& lpszDefExt,
		_In_ const std::wstring& lpszFileName,
		DWORD dwFlags,
		_In_ const std::wstring& lpszFilter,
		_In_opt_ CWnd* pParentWnd)
	{
		auto szFileName = std::wstring(_MAX_PATH, '\0'); // contains full path name after return
		if (dwFlags & OFN_ALLOWMULTISELECT)
		{
			szFileName.resize(CCHBIGBUFF);
		}

		// initialize structure
		auto ofn = OPENFILENAMEEXW{};
		ofn.lStructSize = sizeof(OPENFILENAMEEXW);
		ofn.lpstrFile = const_cast<wchar_t*>(szFileName.c_str());
		ofn.nMaxFile = static_cast<DWORD>(szFileName.length());

		ofn.lpstrDefExt = lpszDefExt.c_str();
		ofn.Flags = dwFlags | OFN_ENABLEHOOK | OFN_EXPLORER;
		ofn.hwndOwner = pParentWnd ? pParentWnd->m_hWnd : nullptr;

		// setup initial file name
		if (!lpszFileName.empty()) wcsncpy_s(ofn.lpstrFile, ofn.nMaxFile, lpszFileName.c_str(), _TRUNCATE);

		// Translate filter into commdlg format (lots of \0)
		auto strFilter = lpszFilter; // filter string
		if (!strFilter.empty())
		{
			replace(strFilter.begin(), strFilter.end(), L'|', L'\0');
			ofn.lpstrFilter = strFilter.c_str();
		}

		BOOL bResult = false;
		if (bOpenFileDialog)
			bResult = GetOpenFileNameW(static_cast<OPENFILENAMEW*>(&ofn));
		else
			bResult = GetSaveFileNameW(static_cast<OPENFILENAMEW*>(&ofn));

		if (bResult)
		{
			m_paths = UnpackFileNames(ofn);
		}

		return bResult ? bResult : IDCANCEL;
	}

	std::wstring CFileDialogExW::GetFileName() const
	{
		if (m_paths.size() >= 1) return m_paths[0];
		return strings::emptystring;
	}

	std::vector<std::wstring> CFileDialogExW::GetFileNames() const { return m_paths; }
} // namespace file