#pragma once
// File.h : header file
//

HRESULT AppendEntryID(LPTSTR szFileName, size_t cchFileName, LPSBinary lpBin, size_t cchMaxAppend);
HRESULT GetDirectoryPath(LPTSTR szPath);
HRESULT SanitizeFileName(LPTSTR szFileOut, size_t cchFileOut, LPCTSTR szFileIn, size_t cchCharsToCopy);

HRESULT BuildFileName(LPTSTR szFileOut, size_t cchFileOut, LPCTSTR szExt, size_t cchExt, LPMESSAGE lpMessage);
HRESULT BuildFileNameAndPath(LPTSTR szFileOut, size_t cchFileOut, LPCTSTR szExt, size_t cchExt, LPCTSTR szSubj, LPSBinary lpBin, LPCTSTR szRootPath);

HRESULT LoadMSGToMessage(LPCTSTR szMessageFile, LPMESSAGE* lppMessage);

HRESULT	LoadFromMSG(LPCTSTR szMessageFile, LPMESSAGE lpMessage, HWND hWnd);
HRESULT	LoadFromTNEF(LPCTSTR szMessageFile, LPADRBOOK lpAdrBook, LPMESSAGE lpMessage);

HRESULT	SaveFolderContentsToMSG(LPMAPIFOLDER lpFolder, LPCTSTR szPathName, BOOL bAssoc, HWND hWnd);
HRESULT	SaveToEML(LPMESSAGE lpMessage, LPCTSTR szFileName);
HRESULT CreateNewMSG(LPCTSTR szFileName, LPMESSAGE* lppMessage, LPSTORAGE* lppStorage);
HRESULT	SaveToMSG(LPMESSAGE lpMessage, LPCTSTR szFileName, HWND hWnd);
HRESULT SaveToTNEF(LPMESSAGE lpMessage, LPADRBOOK lpAdrBook, LPCTSTR szFileName);

HRESULT	DeleteAttachments(LPMESSAGE lpMessage, LPCTSTR szAttName, HWND hWnd);
HRESULT	WriteAttachmentsToFile(LPMESSAGE lpMessage, HWND hWnd);
HRESULT	WriteAttachmentToFile(LPATTACH lpAttach, HWND hWnd);
HRESULT WriteEmbeddedMSGToFile(LPATTACH lpAttach,LPCTSTR szFileName, HWND hWnd);
