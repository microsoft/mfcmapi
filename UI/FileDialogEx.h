#pragma once
// Encapsulate Windows-2000 style open dialog.

class CFileDialogExW
{
public:
	static std::wstring OpenFile(
		_In_ const std::wstring& lpszDefExt,
		_In_ const std::wstring& lpszFileName,
		DWORD dwFlags = 0,
		_In_ const std::wstring& lpszFilter = strings::emptystring,
		_In_opt_ CWnd* pParentWnd = nullptr);

	static std::vector<std::wstring> OpenFiles(
		_In_ const std::wstring& lpszDefExt,
		_In_ const std::wstring& lpszFileName,
		DWORD dwFlags = 0,
		_In_ const std::wstring& lpszFilter = strings::emptystring,
		_In_opt_ CWnd* pParentWnd = nullptr);

	static std::wstring SaveAs(
		_In_ const std::wstring& lpszDefExt,
		_In_ const std::wstring& lpszFileName,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_In_ const std::wstring& lpszFilter = strings::emptystring,
		_In_opt_ CWnd* pParentWnd = nullptr);

private:
	_Check_return_ INT_PTR DisplayDialog(bool bOpenFileDialog, // true for open, false for FileSaveAs
		_In_ const std::wstring& lpszDefExt,
		_In_ const std::wstring& lpszFileName,
		DWORD dwFlags,
		_In_ const std::wstring& lpszFilter = strings::emptystring,
		_In_opt_ CWnd* pParentWnd = nullptr);

	std::wstring GetFileName() const;
	std::vector<std::wstring> GetFileNames() const;

private:
	std::vector<std::wstring> m_paths;
};