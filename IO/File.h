#pragma once
#include <map>

namespace file
{
	std::wstring GetDirectoryPath(HWND hWnd);

#define MAXSUBJ 25
#define MAXSUBJTIGHT 10
#define MAXBIN 141
#define MAXEXT 4
#define MAXATTACH 10
#define MAXMSGPATH (MAX_PATH - MAXSUBJTIGHT - MAXBIN - MAXEXT)
	std::wstring BuildFileName(_In_ const std::wstring& ext, _In_ const std::wstring& dir, _In_ LPMESSAGE lpMessage);
	std::wstring BuildFileNameAndPath(
		_In_ const std::wstring& szExt,
		_In_ const std::wstring& szSubj,
		_In_ const std::wstring& szRootPath,
		_In_opt_ const _SBinary* lpBin);

	_Check_return_ LPMESSAGE LoadMSGToMessage(_In_ const std::wstring& szMessageFile);
	_Check_return_ HRESULT LoadFromMSG(_In_ const std::wstring& szMessageFile, _In_ LPMESSAGE lpMessage, HWND hWnd);
	_Check_return_ HRESULT
	LoadFromTNEF(_In_ const std::wstring& szMessageFile, _In_ LPADRBOOK lpAdrBook, _In_ LPMESSAGE lpMessage);

	void SaveFolderContentsToTXT(
		_In_ LPMDB lpMDB,
		_In_ LPMAPIFOLDER lpFolder,
		bool bRegular,
		bool bAssoc,
		bool bDescend,
		HWND hWnd);
	_Check_return_ HRESULT SaveFolderContentsToMSG(
		_In_ LPMAPIFOLDER lpFolder,
		_In_ const std::wstring& szPathName,
		bool bAssoc,
		bool bUnicode,
		HWND hWnd);
	_Check_return_ HRESULT SaveToEML(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szFileName);
	_Check_return_ HRESULT CreateNewMSG(
		_In_ const std::wstring& szFileName,
		bool bUnicode,
		_Deref_out_opt_ LPMESSAGE* lppMessage,
		_Deref_out_opt_ LPSTORAGE* lppStorage);
	_Check_return_ HRESULT SaveToMSG(
		_In_ LPMAPIFOLDER lpFolder,
		_In_ const std::wstring& szPathName,
		_In_ const SPropValue& entryID,
		_In_ const _SPropValue* lpRecordKey,
		_In_ const _SPropValue* lpSubject,
		bool bUnicode,
		HWND hWnd);
	_Check_return_ HRESULT
	SaveToMSG(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szFileName, bool bUnicode, HWND hWnd, bool bAllowUI);
	_Check_return_ HRESULT
	SaveToTNEF(_In_ LPMESSAGE lpMessage, _In_ LPADRBOOK lpAdrBook, _In_ const std::wstring& szFileName);

	_Check_return_ HRESULT IterateAttachments(
		_In_ LPMESSAGE lpMessage,
		_In_ LPSPropTagArray lpSPropTagArray,
		const std::function<HRESULT(LPSPropValue)>& operation,
		const std::function<bool(HRESULT)>& shouldCancel);
	_Check_return_ HRESULT DeleteAttachments(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szAttName, HWND hWnd);
	_Check_return_ HRESULT
	WriteEmbeddedMSGToFile(_In_ LPATTACH lpAttach, _In_ const std::wstring& szFileName, bool bUnicode, HWND hWnd);
	_Check_return_ HRESULT WriteAttachStreamToFile(_In_ LPATTACH lpAttach, _In_ const std::wstring& szFileName);
	_Check_return_ HRESULT WriteOleToFile(_In_ LPATTACH lpAttach, _In_ const std::wstring& szFileName);

	std::wstring GetModuleFileName(_In_opt_ HMODULE hModule);
	std::wstring GetSystemDirectory();
	std::map<std::wstring, std::wstring> GetFileVersionInfo(_In_opt_ HMODULE hModule);
} // namespace file