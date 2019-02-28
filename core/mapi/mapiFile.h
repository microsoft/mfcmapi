#pragma once

namespace file
{
	std::wstring BuildFileName(_In_ const std::wstring& ext, _In_ const std::wstring& dir, _In_ LPMESSAGE lpMessage);

	_Check_return_ LPMESSAGE LoadMSGToMessage(_In_ const std::wstring& szMessageFile);
	_Check_return_ HRESULT LoadFromMSG(_In_ const std::wstring& szMessageFile, _In_ LPMESSAGE lpMessage, HWND hWnd);
	_Check_return_ HRESULT
	LoadFromTNEF(_In_ const std::wstring& szMessageFile, _In_ LPADRBOOK lpAdrBook, _In_ LPMESSAGE lpMessage);

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

	_Check_return_ STDMETHODIMP MyOpenStreamOnFile(
		_In_ LPALLOCATEBUFFER lpAllocateBuffer,
		_In_ LPFREEBUFFER lpFreeBuffer,
		ULONG ulFlags,
		_In_ const std::wstring& lpszFileName,
		_Out_ LPSTREAM* lppStream);
} // namespace file