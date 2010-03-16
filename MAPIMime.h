#pragma once

// Constants - http://msdn2.microsoft.com/en-us/library/bb905201.aspx
#define CCSF_SMTP             0x0002 // the converter is being passed an SMTP message
#define CCSF_NOHEADERS        0x0004 // the converter should ignore the headers on the outside message
#define CCSF_USE_TNEF         0x0010 // the converter should embed TNEF in the MIME message
#define CCSF_INCLUDE_BCC      0x0020 // the converter should include Bcc recipients
#define CCSF_8BITHEADERS      0x0040 // the converter should allow 8 bit headers
#define CCSF_USE_RTF          0x0080 // the converter should do HTML->RTF conversion
#define CCSF_PLAIN_TEXT_ONLY  0x1000 // the converter should just send plain text
#define CCSF_NO_MSGID         0x4000 // don't include Message-Id field in outgoing messages
#define CCSF_EMBEDDED_MESSAGE 0x8000 // sent/unsent information is persisted in X-Unsent

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

	virtual HRESULT STDMETHODCALLTYPE SetTextWrapping(BOOL  fWrapText,
		ULONG ulWrapWidth);

	virtual HRESULT STDMETHODCALLTYPE SetSaveFormat(MIMESAVETYPE mstSaveFormat);

	virtual HRESULT PlaceHolder5();

	virtual HRESULT STDMETHODCALLTYPE SetCharset(
		BOOL fApply,
		HCHARSET hcharset,
		CSETAPPLYTYPE csetapplytype);
};

typedef IConverterSession FAR * LPCONVERTERSESSION;

// Helper functions
HRESULT ImportEMLToIMessage(
							LPCWSTR lpszEMLFile,
							LPMESSAGE lpMsg,
							ULONG ulConvertFlags,
							BOOL bApply,
							HCHARSET hCharSet,
							CSETAPPLYTYPE cSetApplyType,
							LPADRBOOK lpAdrBook);
HRESULT ExportIMessageToEML(LPMESSAGE lpMsg, LPCWSTR lpszEMLFile, ULONG ulConvertFlags,
							ENCODINGTYPE et, MIMESAVETYPE mst, ULONG ulWrapLines,LPADRBOOK lpAdrBook);
HRESULT ConvertEMLToMSG(
						LPCWSTR lpszEMLFile,
						LPCWSTR lpszMSGFile,
						ULONG ulConvertFlags,
						BOOL bApply,
						HCHARSET hCharSet,
						CSETAPPLYTYPE cSetApplyType,
						LPADRBOOK lpAdrBook,
						BOOL bUnicode);
HRESULT ConvertMSGToEML(LPCWSTR lpszMSGFile, LPCWSTR lpszEMLFile, ULONG ulConvertFlags,
						ENCODINGTYPE et, MIMESAVETYPE mst, ULONG ulWrapLines,
						LPADRBOOK lpAdrBook);
HRESULT GetConversionToEMLOptions(CWnd* pParentWnd,
								  ULONG* lpulConvertFlags,
								  ENCODINGTYPE* lpet,
								  MIMESAVETYPE* lpmst,
								  ULONG* lpulWrapLines,
								  BOOL* pDoAdrBook);
HRESULT GetConversionFromEMLOptions(CWnd* pParentWnd,
									ULONG* lpulConvertFlags,
									BOOL* pDoAdrBook,
									BOOL* pDoApply,
									HCHARSET* phCharSet,
									CSETAPPLYTYPE* pcSetApplyType,
									BOOL* pbUnicode);
