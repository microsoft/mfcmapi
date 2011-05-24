#pragma once

#define USE_DEFAULT_WRAPPING 0xFFFFFFFF
#define USE_DEFAULT_SAVETYPE (MIMESAVETYPE) 0xFFFFFFFF

const ULARGE_INTEGER ULARGE_MAX = {0xFFFFFFFFU, 0xFFFFFFFFU};

// http://msdn2.microsoft.com/en-us/library/bb905202.aspx
interface IConverterSession : public IUnknown
{
public:
	virtual HRESULT STDMETHODCALLTYPE SetAdrBook(LPADRBOOK pab);

	virtual HRESULT STDMETHODCALLTYPE SetEncoding(ENCODINGTYPE et);

	virtual HRESULT PlaceHolder1();

	virtual HRESULT STDMETHODCALLTYPE MIMEToMAPI(LPSTREAM pstm,
		LPMESSAGE pmsg,
		LPCSTR pszSrcSrv,
		ULONG ulFlags);

	virtual HRESULT STDMETHODCALLTYPE MAPIToMIMEStm(LPMESSAGE pmsg,
		LPSTREAM pstm,
		ULONG ulFlags);

	virtual HRESULT PlaceHolder2();
	virtual HRESULT PlaceHolder3();
	virtual HRESULT PlaceHolder4();

	virtual HRESULT STDMETHODCALLTYPE SetTextWrapping(bool  fWrapText,
		ULONG ulWrapWidth);

	virtual HRESULT STDMETHODCALLTYPE SetSaveFormat(MIMESAVETYPE mstSaveFormat);

	virtual HRESULT PlaceHolder5();

	virtual HRESULT STDMETHODCALLTYPE SetCharset(
		bool fApply,
		HCHARSET hcharset,
		CSETAPPLYTYPE csetapplytype);
};

typedef IConverterSession* LPCONVERTERSESSION;

// Helper functions
_Check_return_ HRESULT ImportEMLToIMessage(
	_In_z_ LPCWSTR lpszEMLFile,
	_In_ LPMESSAGE lpMsg,
	ULONG ulConvertFlags,
	bool bApply,
	HCHARSET hCharSet,
	CSETAPPLYTYPE cSetApplyType,
	_In_opt_ LPADRBOOK lpAdrBook);
_Check_return_ HRESULT ExportIMessageToEML(_In_ LPMESSAGE lpMsg, _In_z_ LPCWSTR lpszEMLFile, ULONG ulConvertFlags,
										   ENCODINGTYPE et, MIMESAVETYPE mst, ULONG ulWrapLines, _In_opt_ LPADRBOOK lpAdrBook);
_Check_return_ HRESULT ConvertEMLToMSG(_In_z_ LPCWSTR lpszEMLFile,
									   _In_z_ LPCWSTR lpszMSGFile,
									   ULONG ulConvertFlags,
									   bool bApply,
									   HCHARSET hCharSet,
									   CSETAPPLYTYPE cSetApplyType,
									   _In_opt_ LPADRBOOK lpAdrBook,
									   bool bUnicode);
_Check_return_ HRESULT ConvertMSGToEML(_In_z_ LPCWSTR lpszMSGFile, _In_z_ LPCWSTR lpszEMLFile, ULONG ulConvertFlags,
									   ENCODINGTYPE et, MIMESAVETYPE mst, ULONG ulWrapLines,
									   _In_opt_ LPADRBOOK lpAdrBook);
_Check_return_ HRESULT GetConversionToEMLOptions(_In_ CWnd* pParentWnd,
												 _Out_ ULONG* lpulConvertFlags,
												 _Out_ ENCODINGTYPE* lpet,
												 _Out_ MIMESAVETYPE* lpmst,
												 _Out_ ULONG* lpulWrapLines,
												 _Out_ bool* pDoAdrBook);
_Check_return_ HRESULT GetConversionFromEMLOptions(_In_ CWnd* pParentWnd,
												   _Out_ ULONG* lpulConvertFlags,
												   _Out_ bool* pDoAdrBook,
												   _Out_ bool* pDoApply,
												   _Out_ HCHARSET* phCharSet,
												   _Out_ CSETAPPLYTYPE* pcSetApplyType,
												   _Out_opt_ bool* pbUnicode);