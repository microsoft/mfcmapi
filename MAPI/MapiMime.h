#pragma once
#include <Interpret/ExtraPropTags.h>

namespace mapi
{
	namespace mapimime
	{
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

			virtual HRESULT STDMETHODCALLTYPE MIMEToMAPI(
				LPSTREAM pstm, LPMESSAGE pmsg, LPCSTR pszSrcSrv, ULONG ulFlags);

			virtual HRESULT STDMETHODCALLTYPE MAPIToMIMEStm(LPMESSAGE pmsg, LPSTREAM pstm, ULONG ulFlags);

			virtual HRESULT PlaceHolder2();
			virtual HRESULT PlaceHolder3();
			virtual HRESULT PlaceHolder4();

			virtual HRESULT STDMETHODCALLTYPE SetTextWrapping(bool fWrapText, ULONG ulWrapWidth);

			virtual HRESULT STDMETHODCALLTYPE SetSaveFormat(MIMESAVETYPE mstSaveFormat);

			virtual HRESULT PlaceHolder5();

			virtual HRESULT STDMETHODCALLTYPE SetCharset(bool fApply, HCHARSET hcharset, CSETAPPLYTYPE csetapplytype);
		};

		typedef IConverterSession* LPCONVERTERSESSION;

		// Helper functions
		_Check_return_ HRESULT ImportEMLToIMessage(
			_In_z_ LPCWSTR lpszEMLFile,
			_In_ LPMESSAGE lpMsg,
			CCSFLAGS convertFlags,
			bool bApply,
			HCHARSET hCharSet,
			CSETAPPLYTYPE cSetApplyType,
			_In_opt_ LPADRBOOK lpAdrBook);
		_Check_return_ HRESULT ExportIMessageToEML(
			_In_ LPMESSAGE lpMsg,
			_In_z_ LPCWSTR lpszEMLFile,
			CCSFLAGS convertFlags,
			ENCODINGTYPE et,
			MIMESAVETYPE mst,
			ULONG ulWrapLines,
			_In_opt_ LPADRBOOK lpAdrBook);
		_Check_return_ HRESULT ConvertEMLToMSG(
			_In_z_ LPCWSTR lpszEMLFile,
			_In_z_ LPCWSTR lpszMSGFile,
			CCSFLAGS convertFlags,
			bool bApply,
			HCHARSET hCharSet,
			CSETAPPLYTYPE cSetApplyType,
			_In_opt_ LPADRBOOK lpAdrBook,
			bool bUnicode);
		_Check_return_ HRESULT ConvertMSGToEML(
			_In_z_ LPCWSTR lpszMSGFile,
			_In_z_ LPCWSTR lpszEMLFile,
			CCSFLAGS convertFlags,
			ENCODINGTYPE et,
			MIMESAVETYPE mst,
			ULONG ulWrapLines,
			_In_opt_ LPADRBOOK lpAdrBook);
	} // namespace mapimime
} // namespace mapi
