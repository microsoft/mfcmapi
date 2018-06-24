#include <StdAfx.h>
#include <MAPI/MapiMime.h>
#include <IO/File.h>
#include <Interpret/Guids.h>
#include <ImportProcs.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <MAPI/MAPIFunctions.h>
#ifndef MRMAPI
#include <Interpret/ExtraPropTags.h>
#endif

namespace mapi
{
	namespace mapimime
	{
		_Check_return_ HRESULT ImportEMLToIMessage(
			_In_z_ LPCWSTR lpszEMLFile,
			_In_ LPMESSAGE lpMsg,
			ULONG ulConvertFlags,
			bool bApply,
			HCHARSET hCharSet,
			CSETAPPLYTYPE cSetApplyType,
			_In_opt_ LPADRBOOK lpAdrBook)
		{
			if (!lpszEMLFile || !lpMsg) return MAPI_E_INVALID_PARAMETER;

			LPCONVERTERSESSION lpConverter = nullptr;

			auto hRes = EC_H_MSG(
				IDS_NOCONVERTERSESSION,
				CoCreateInstance(
					guid::CLSID_IConverterSession,
					nullptr,
					CLSCTX_INPROC_SERVER,
					guid::IID_IConverterSession,
					reinterpret_cast<LPVOID*>(&lpConverter)));
			if (SUCCEEDED(hRes) && lpConverter)
			{
				LPSTREAM lpEMLStm = nullptr;

				EC_H(mapi::MyOpenStreamOnFile(MAPIAllocateBuffer, MAPIFreeBuffer, STGM_READ, lpszEMLFile, &lpEMLStm));
				if (SUCCEEDED(hRes) && lpEMLStm)
				{
					if (lpAdrBook)
					{
						EC_MAPI(lpConverter->SetAdrBook(lpAdrBook));
					}
					if (SUCCEEDED(hRes) && bApply)
					{
						EC_MAPI(lpConverter->SetCharset(bApply, hCharSet, cSetApplyType));
					}
					if (SUCCEEDED(hRes))
					{
						// We'll make the user ensure CCSF_SMTP is passed
						EC_MAPI(lpConverter->MIMEToMAPI(
							lpEMLStm,
							lpMsg,
							nullptr, // Must be nullptr
							ulConvertFlags));
						if (SUCCEEDED(hRes))
						{
							EC_MAPI(lpMsg->SaveChanges(NULL));
						}
					}
				}

				if (lpEMLStm) lpEMLStm->Release();
			}

			if (lpConverter) lpConverter->Release();

			return hRes;
		}

		_Check_return_ HRESULT ExportIMessageToEML(
			_In_ LPMESSAGE lpMsg,
			_In_z_ LPCWSTR lpszEMLFile,
			ULONG ulConvertFlags,
			ENCODINGTYPE et,
			MIMESAVETYPE mst,
			ULONG ulWrapLines,
			_In_opt_ LPADRBOOK lpAdrBook)
		{
			if (!lpszEMLFile || !lpMsg) return MAPI_E_INVALID_PARAMETER;

			LPCONVERTERSESSION lpConverter = nullptr;

			auto hRes = EC_H_MSG(
				IDS_NOCONVERTERSESSION,
				CoCreateInstance(
					guid::CLSID_IConverterSession,
					nullptr,
					CLSCTX_INPROC_SERVER,
					guid::IID_IConverterSession,
					reinterpret_cast<LPVOID*>(&lpConverter)));
			if (SUCCEEDED(hRes) && lpConverter)
			{
				if (lpAdrBook)
				{
					hRes = EC_MAPI2(lpConverter->SetAdrBook(lpAdrBook));
				}

				if (SUCCEEDED(hRes) && et != IET_UNKNOWN)
				{
					hRes = EC_MAPI2(lpConverter->SetEncoding(et));
				}

				if (SUCCEEDED(hRes) && mst != USE_DEFAULT_SAVETYPE)
				{
					hRes = EC_MAPI2(lpConverter->SetSaveFormat(mst));
				}

				if (SUCCEEDED(hRes) && ulWrapLines != USE_DEFAULT_WRAPPING)
				{
					hRes = EC_MAPI2(lpConverter->SetTextWrapping(ulWrapLines != 0, ulWrapLines));
				}

				if (SUCCEEDED(hRes))
				{
					LPSTREAM lpMimeStm = nullptr;

					hRes = EC_H2(CreateStreamOnHGlobal(nullptr, true, &lpMimeStm));
					if (SUCCEEDED(hRes) && lpMimeStm)
					{
						// Per the docs for MAPIToMIMEStm, CCSF_SMTP must always be set
						// But we're gonna make the user ensure that, so we don't or it in here
						hRes = EC_MAPI2(lpConverter->MAPIToMIMEStm(lpMsg, lpMimeStm, ulConvertFlags));
						if (SUCCEEDED(hRes))
						{
							LPSTREAM lpFileStm = nullptr;

							hRes = EC_H2(mapi::MyOpenStreamOnFile(
								MAPIAllocateBuffer,
								MAPIFreeBuffer,
								STGM_CREATE | STGM_READWRITE,
								lpszEMLFile,
								&lpFileStm));
							if (SUCCEEDED(hRes) && lpFileStm)
							{
								const LARGE_INTEGER dwBegin = {0};
								hRes = EC_MAPI2(lpMimeStm->Seek(dwBegin, STREAM_SEEK_SET, nullptr));
								if (SUCCEEDED(hRes))
								{
									hRes = EC_MAPI2(lpMimeStm->CopyTo(lpFileStm, ULARGE_MAX, nullptr, nullptr));
								}

								if (SUCCEEDED(hRes))
								{
									hRes = EC_MAPI2(lpFileStm->Commit(STGC_DEFAULT));
								}
							}

							if (lpFileStm) lpFileStm->Release();
						}
					}

					if (lpMimeStm) lpMimeStm->Release();
				}
			}

			if (lpConverter) lpConverter->Release();

			return hRes;
		}

		_Check_return_ HRESULT ConvertEMLToMSG(
			_In_z_ LPCWSTR lpszEMLFile,
			_In_z_ LPCWSTR lpszMSGFile,
			ULONG ulConvertFlags,
			bool bApply,
			HCHARSET hCharSet,
			CSETAPPLYTYPE cSetApplyType,
			_In_opt_ LPADRBOOK lpAdrBook,
			bool bUnicode)
		{
			if (!lpszEMLFile || !lpszMSGFile) return MAPI_E_INVALID_PARAMETER;

			auto hRes = S_OK;
			LPSTORAGE pStorage = nullptr;
			LPMESSAGE pMessage = nullptr;

			EC_H(file::CreateNewMSG(lpszMSGFile, bUnicode, &pMessage, &pStorage));
			if (SUCCEEDED(hRes) && pMessage && pStorage)
			{
				EC_H(ImportEMLToIMessage(
					lpszEMLFile, pMessage, ulConvertFlags, bApply, hCharSet, cSetApplyType, lpAdrBook));
				if (SUCCEEDED(hRes))
				{
					EC_MAPI(pStorage->Commit(STGC_DEFAULT));
				}
			}

			if (pMessage) pMessage->Release();
			if (pStorage) pStorage->Release();

			return hRes;
		}

		_Check_return_ HRESULT ConvertMSGToEML(
			_In_z_ LPCWSTR lpszMSGFile,
			_In_z_ LPCWSTR lpszEMLFile,
			ULONG ulConvertFlags,
			ENCODINGTYPE et,
			MIMESAVETYPE mst,
			ULONG ulWrapLines,
			_In_opt_ LPADRBOOK lpAdrBook)
		{
			if (!lpszEMLFile || !lpszMSGFile) return MAPI_E_INVALID_PARAMETER;

			auto hRes = S_OK;
			LPMESSAGE pMessage = nullptr;

			EC_H(file::LoadMSGToMessage(lpszMSGFile, &pMessage));
			if (SUCCEEDED(hRes) && pMessage)
			{
				EC_H(ExportIMessageToEML(pMessage, lpszEMLFile, ulConvertFlags, et, mst, ulWrapLines, lpAdrBook));
			}

			if (pMessage) pMessage->Release();

			return hRes;
		}

#ifndef MRMAPI
		_Check_return_ HRESULT GetConversionToEMLOptions(
			_In_ CWnd* pParentWnd,
			_Out_ ULONG* lpulConvertFlags,
			_Out_ ENCODINGTYPE* lpet,
			_Out_ MIMESAVETYPE* lpmst,
			_Out_ ULONG* lpulWrapLines,
			_Out_ bool* pDoAdrBook)
		{
			if (!lpulConvertFlags || !lpet || !lpmst || !lpulWrapLines || !pDoAdrBook) return MAPI_E_INVALID_PARAMETER;
			auto hRes = S_OK;

			dialog::editor::CEditor MyData(
				pParentWnd, IDS_CONVERTTOEML, IDS_CONVERTTOEMLPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_CONVERTFLAGS, false));
			MyData.SetHex(0, CCSF_SMTP);
			MyData.InitPane(1, viewpane::CheckPane::Create(IDS_CONVERTDOENCODINGTYPE, false, false));
			MyData.InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_CONVERTENCODINGTYPE, false));
			MyData.SetHex(2, IET_7BIT);
			MyData.InitPane(3, viewpane::CheckPane::Create(IDS_CONVERTDOMIMESAVETYPE, false, false));
			MyData.InitPane(4, viewpane::TextPane::CreateSingleLinePane(IDS_CONVERTMIMESAVETYPE, false));
			MyData.SetHex(4, SAVE_RFC822);
			MyData.InitPane(5, viewpane::CheckPane::Create(IDS_CONVERTDOWRAPLINES, false, false));
			MyData.InitPane(6, viewpane::TextPane::CreateSingleLinePane(IDS_CONVERTWRAPLINECOUNT, false));
			MyData.SetDecimal(6, 74);
			MyData.InitPane(7, viewpane::CheckPane::Create(IDS_CONVERTDOADRBOOK, false, false));

			WC_H(MyData.DisplayDialog());
			if (hRes == S_OK)
			{
				*lpulConvertFlags = MyData.GetHex(0);
				*lpulWrapLines = MyData.GetCheck(1) ? static_cast<ENCODINGTYPE>(MyData.GetDecimal(2)) : IET_UNKNOWN;
				*lpmst = MyData.GetCheck(3) ? static_cast<MIMESAVETYPE>(MyData.GetHex(4)) : USE_DEFAULT_SAVETYPE;
				*lpulWrapLines = MyData.GetCheck(5) ? MyData.GetDecimal(6) : USE_DEFAULT_WRAPPING;
				*pDoAdrBook = MyData.GetCheck(7);
			}
			return hRes;
		}

		_Check_return_ HRESULT GetConversionFromEMLOptions(
			_In_ CWnd* pParentWnd,
			_Out_ ULONG* lpulConvertFlags,
			_Out_ bool* pDoAdrBook,
			_Out_ bool* pDoApply,
			_Out_ HCHARSET* phCharSet,
			_Out_ CSETAPPLYTYPE* pcSetApplyType,
			_Out_opt_ bool* pbUnicode)
		{
			if (!lpulConvertFlags || !pDoAdrBook || !pDoApply || !phCharSet || !pcSetApplyType)
				return MAPI_E_INVALID_PARAMETER;
			auto hRes = S_OK;

			dialog::editor::CEditor MyData(
				pParentWnd, IDS_CONVERTFROMEML, IDS_CONVERTFROMEMLPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_CONVERTFLAGS, false));
			MyData.SetHex(0, CCSF_SMTP);
			MyData.InitPane(1, viewpane::CheckPane::Create(IDS_CONVERTCODEPAGE, false, false));
			MyData.InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_CONVERTCODEPAGE, false));
			MyData.SetDecimal(2, CP_USASCII);
			MyData.InitPane(3, viewpane::TextPane::CreateSingleLinePane(IDS_CONVERTCHARSETTYPE, false));
			MyData.SetDecimal(3, CHARSET_BODY);
			MyData.InitPane(4, viewpane::TextPane::CreateSingleLinePane(IDS_CONVERTCHARSETAPPLYTYPE, false));
			MyData.SetDecimal(4, CSET_APPLY_UNTAGGED);
			MyData.InitPane(5, viewpane::CheckPane::Create(IDS_CONVERTDOADRBOOK, false, false));
			if (pbUnicode)
			{
				MyData.InitPane(6, viewpane::CheckPane::Create(IDS_SAVEUNICODE, false, false));
			}

			WC_H(MyData.DisplayDialog());
			if (hRes == S_OK)
			{
				*lpulConvertFlags = MyData.GetHex(0);
				if (MyData.GetCheck(1))
				{
					if (SUCCEEDED(hRes)) *pDoApply = true;
					*pcSetApplyType = static_cast<CSETAPPLYTYPE>(MyData.GetDecimal(4));
					if (*pcSetApplyType > CSET_APPLY_TAG_ALL) return MAPI_E_INVALID_PARAMETER;
					const auto ulCodePage = MyData.GetDecimal(2);
					const auto cCharSetType = static_cast<CHARSETTYPE>(MyData.GetDecimal(3));
					if (cCharSetType > CHARSET_WEB) return MAPI_E_INVALID_PARAMETER;
					EC_H(import::MyMimeOleGetCodePageCharset(ulCodePage, cCharSetType, phCharSet));
				}
				*pDoAdrBook = MyData.GetCheck(5);
				if (pbUnicode)
				{
					*pbUnicode = MyData.GetCheck(6);
				}
			}
			return hRes;
		}
#endif
	}
}