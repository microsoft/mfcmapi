#pragma once
// File.h : header file
//

HRESULT AppendEntryID(LPWSTR szFileName, size_t cchFileName, LPSBinary lpBin, size_t cchMaxAppend);
HRESULT GetDirectoryPath(LPWSTR szPath);
HRESULT SanitizeFileNameA(LPSTR szFileOut, size_t cchFileOut, LPCSTR szFileIn, size_t cchCharsToCopy);
HRESULT SanitizeFileNameW(LPWSTR szFileOut, size_t cchFileOut, LPCWSTR szFileIn, size_t cchCharsToCopy);
#ifdef UNICODE
#define SanitizeFileName SanitizeFileNameW
#else
#define SanitizeFileName SanitizeFileNameA
#endif

HRESULT BuildFileName(LPWSTR szFileOut, size_t cchFileOut, LPCWSTR szExt, size_t cchExt, LPMESSAGE lpMessage);
HRESULT BuildFileNameAndPath(LPWSTR szFileOut, size_t cchFileOut, LPCWSTR szExt, size_t cchExt, LPCWSTR szSubj, LPSBinary lpBin, LPCWSTR szRootPath);

HRESULT LoadMSGToMessage(LPCWSTR szMessageFile, LPMESSAGE* lppMessage);

HRESULT	LoadFromMSG(LPCWSTR szMessageFile, LPMESSAGE lpMessage, HWND hWnd);
HRESULT	LoadFromTNEF(LPCWSTR szMessageFile, LPADRBOOK lpAdrBook, LPMESSAGE lpMessage);

HRESULT	SaveFolderContentsToMSG(LPMAPIFOLDER lpFolder, LPCWSTR szPathName, BOOL bAssoc, BOOL bUnicode, HWND hWnd);
HRESULT	SaveToEML(LPMESSAGE lpMessage, LPCWSTR szFileName);
HRESULT CreateNewMSG(LPCWSTR szFileName, BOOL bUnicode, LPMESSAGE* lppMessage, LPSTORAGE* lppStorage);
HRESULT	SaveToMSG(LPMESSAGE lpMessage, LPCWSTR szFileName, BOOL bUnicode, HWND hWnd);
HRESULT SaveToTNEF(LPMESSAGE lpMessage, LPADRBOOK lpAdrBook, LPCWSTR szFileName);

HRESULT	DeleteAttachments(LPMESSAGE lpMessage, LPCTSTR szAttName, HWND hWnd);
HRESULT	WriteAttachmentsToFile(LPMESSAGE lpMessage, HWND hWnd);
HRESULT	WriteAttachmentToFile(LPATTACH lpAttach, HWND hWnd);
HRESULT WriteEmbeddedMSGToFile(LPATTACH lpAttach,LPCWSTR szFileName, BOOL bUnicode, HWND hWnd);
