#pragma once
// File.h : header file

_Check_return_ HRESULT GetDirectoryPath(HWND hWnd, _Inout_z_ LPWSTR szPath);
_Check_return_ HRESULT SanitizeFileNameA(_Inout_z_count_(cchFileOut) LPSTR szFileOut, size_t cchFileOut, _In_z_ LPCSTR szFileIn, size_t cchCharsToCopy);
_Check_return_ HRESULT SanitizeFileNameW(_Inout_z_count_(cchFileOut) LPWSTR szFileOut, size_t cchFileOut, _In_z_ LPCWSTR szFileIn, size_t cchCharsToCopy);
wstring SanitizeFileNameW(_In_ wstring szFileIn);
#ifdef UNICODE
#define SanitizeFileName SanitizeFileNameW
#else
#define SanitizeFileName SanitizeFileNameA
#endif

wstring BuildFileName(
	_In_ wstring szExt,
	_In_ LPMESSAGE lpMessage);
wstring BuildFileNameAndPath(
	_In_ wstring szExt,
	_In_ wstring szSubj,
	_In_ wstring szRootPath,
	_In_opt_ LPSBinary lpBin);

_Check_return_ HRESULT LoadMSGToMessage(_In_z_ LPCWSTR szMessageFile, _Deref_out_opt_ LPMESSAGE* lppMessage);

_Check_return_ HRESULT LoadFromMSG(_In_z_ LPCWSTR szMessageFile, _In_ LPMESSAGE lpMessage, HWND hWnd);
_Check_return_ HRESULT LoadFromTNEF(_In_z_ LPCWSTR szMessageFile, _In_ LPADRBOOK lpAdrBook, _In_ LPMESSAGE lpMessage);

void SaveFolderContentsToTXT(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpFolder, bool bRegular, bool bAssoc, bool bDescend, HWND hWnd);
_Check_return_ HRESULT SaveFolderContentsToMSG(_In_ LPMAPIFOLDER lpFolder, _In_z_ LPCWSTR szPathName, bool bAssoc, bool bUnicode, HWND hWnd);
_Check_return_ HRESULT SaveToEML(_In_ LPMESSAGE lpMessage, _In_z_ LPCWSTR szFileName);
_Check_return_ HRESULT CreateNewMSG(_In_z_ LPCWSTR szFileName, bool bUnicode, _Deref_out_opt_ LPMESSAGE* lppMessage, _Deref_out_opt_ LPSTORAGE* lppStorage);
_Check_return_ HRESULT SaveToMSG(_In_ LPMESSAGE lpMessage, _In_z_ LPCWSTR szFileName, bool bUnicode, HWND hWnd, bool bAllowUI);
_Check_return_ HRESULT SaveToTNEF(_In_ LPMESSAGE lpMessage, _In_ LPADRBOOK lpAdrBook, _In_z_ LPCWSTR szFileName);

_Check_return_ HRESULT DeleteAttachments(_In_ LPMESSAGE lpMessage, _In_ wstring& szAttName, HWND hWnd);
_Check_return_ HRESULT WriteAttachmentsToFile(_In_ LPMESSAGE lpMessage, HWND hWnd);
_Check_return_ HRESULT WriteAttachmentToFile(_In_ LPATTACH lpAttach, HWND hWnd);
_Check_return_ HRESULT WriteEmbeddedMSGToFile(_In_ LPATTACH lpAttach, _In_z_ LPCWSTR szFileName, bool bUnicode, HWND hWnd);