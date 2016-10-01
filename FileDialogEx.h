#pragma once
// CFileDialogExW: Encapsulate Windows-2000 style open dialog.

class CFileDialogExW
{
public:
	static wstring OpenFile(
		_In_ wstring lpszDefExt,
		_In_ wstring lpszFileName,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_In_ wstring lpszFilter = emptystring,
		_In_opt_ CWnd* pParentWnd = nullptr);

	static vector<wstring> OpenFiles(
		_In_ wstring lpszDefExt,
		_In_ wstring lpszFileName,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_In_ wstring lpszFilter = emptystring,
		_In_opt_ CWnd* pParentWnd = nullptr);

	static wstring SaveAs(
		_In_ wstring lpszDefExt,
		_In_ wstring lpszFileName,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_In_ wstring lpszFilter = emptystring,
		_In_opt_ CWnd* pParentWnd = nullptr);

private:
	_Check_return_ INT_PTR DisplayDialog(bool bOpenFileDialog, // true for open, false for FileSaveAs
		_In_ wstring lpszDefExt,
		_In_ wstring lpszFileName,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		_In_ wstring lpszFilter = emptystring,
		_In_opt_ CWnd* pParentWnd = nullptr);

	wstring GetFileName() const;
	vector<wstring> GetFileNames() const;

private:
	vector<wstring> m_paths;
};