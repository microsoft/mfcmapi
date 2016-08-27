#pragma once
// File.h : header file
//

_Check_return_ HRESULT AppendEntryID(_Inout_z_count_(cchFileName) LPWSTR szFileName, size_t cchFileName, _In_ LPSBinary lpBin, size_t cchMaxAppend);
_Check_return_ HRESULT GetDirectoryPath(HWND hWnd, _Inout_z_ LPWSTR szPath);
_Check_return_ HRESULT SanitizeFileNameA(_Inout_z_count_(cchFileOut) LPSTR szFileOut, size_t cchFileOut, _In_z_ LPCSTR szFileIn, size_t cchCharsToCopy);
_Check_return_ HRESULT SanitizeFileNameW(_Inout_z_count_(cchFileOut) LPWSTR szFileOut, size_t cchFileOut, _In_z_ LPCWSTR szFileIn, size_t cchCharsToCopy);
#ifdef UNICODE
#define SanitizeFileName SanitizeFileNameW
#else
#define SanitizeFileName SanitizeFileNameA
#endif

_Check_return_ HRESULT BuildFileName(_Inout_z_count_(cchFileOut) LPWSTR szFileOut,
									 size_t cchFileOut,
									 _In_z_count_(cchExt) LPCWSTR szExt,
									 size_t cchExt,
									 _In_ LPMESSAGE lpMessage);
_Check_return_ HRESULT BuildFileNameAndPath(_Inout_z_count_(cchFileOut) LPWSTR szFileOut,
											size_t cchFileOut,
											_In_z_count_(cchExt) LPCWSTR szExt,
											size_t cchExt,
											_In_opt_z_ LPCWSTR szSubj,
											_In_opt_ LPSBinary lpBin,
											_In_opt_z_ LPCWSTR szRootPath);

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
_Check_return_ HRESULT WriteEmbeddedMSGToFile(_In_ LPATTACH lpAttach,_In_z_ LPCWSTR szFileName, bool bUnicode, HWND hWnd);
