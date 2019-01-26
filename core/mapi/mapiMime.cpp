#include <core/stdafx.h>
#include <core/mapi/mapiMime.h>
#include <core/interpret/guid.h>
#include <core/mapi/mapiFile.h>
#include <core/utility/error.h>
#include <core/mapi/extraPropTags.h>

namespace mapi
{
	namespace mapimime
	{
		_Check_return_ HRESULT ImportEMLToIMessage(
			_In_z_ LPCWSTR lpszEMLFile,
			_In_ LPMESSAGE lpMsg,
			CCSFLAGS convertFlags,
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

				hRes = EC_H(
					file::MyOpenStreamOnFile(MAPIAllocateBuffer, MAPIFreeBuffer, STGM_READ, lpszEMLFile, &lpEMLStm));
				if (SUCCEEDED(hRes) && lpEMLStm)
				{
					if (lpAdrBook)
					{
						hRes = EC_MAPI(lpConverter->SetAdrBook(lpAdrBook));
					}

					if (SUCCEEDED(hRes) && bApply)
					{
						hRes = EC_MAPI(lpConverter->SetCharset(bApply, hCharSet, cSetApplyType));
					}

					if (SUCCEEDED(hRes))
					{
						// We'll make the user ensure CCSF_SMTP is passed
						hRes = EC_MAPI(lpConverter->MIMEToMAPI(
							lpEMLStm,
							lpMsg,
							nullptr, // Must be nullptr
							convertFlags));
						if (SUCCEEDED(hRes))
						{
							hRes = EC_MAPI(lpMsg->SaveChanges(NULL));
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
			CCSFLAGS convertFlags,
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
					hRes = EC_MAPI(lpConverter->SetAdrBook(lpAdrBook));
				}

				if (SUCCEEDED(hRes) && et != IET_UNKNOWN)
				{
					hRes = EC_MAPI(lpConverter->SetEncoding(et));
				}

				if (SUCCEEDED(hRes) && mst != USE_DEFAULT_SAVETYPE)
				{
					hRes = EC_MAPI(lpConverter->SetSaveFormat(mst));
				}

				if (SUCCEEDED(hRes) && ulWrapLines != USE_DEFAULT_WRAPPING)
				{
					hRes = EC_MAPI(lpConverter->SetTextWrapping(ulWrapLines != 0, ulWrapLines));
				}

				if (SUCCEEDED(hRes))
				{
					LPSTREAM lpMimeStm = nullptr;

					hRes = EC_H(CreateStreamOnHGlobal(nullptr, true, &lpMimeStm));
					if (SUCCEEDED(hRes) && lpMimeStm)
					{
						// Per the docs for MAPIToMIMEStm, CCSF_SMTP must always be set
						// But we're gonna make the user ensure that, so we don't or it in here
						hRes = EC_MAPI(lpConverter->MAPIToMIMEStm(lpMsg, lpMimeStm, convertFlags));
						if (SUCCEEDED(hRes))
						{
							LPSTREAM lpFileStm = nullptr;

							hRes = EC_H(file::MyOpenStreamOnFile(
								MAPIAllocateBuffer,
								MAPIFreeBuffer,
								STGM_CREATE | STGM_READWRITE,
								lpszEMLFile,
								&lpFileStm));
							if (SUCCEEDED(hRes) && lpFileStm)
							{
								const LARGE_INTEGER dwBegin = {};
								hRes = EC_MAPI(lpMimeStm->Seek(dwBegin, STREAM_SEEK_SET, nullptr));
								if (SUCCEEDED(hRes))
								{
									hRes = EC_MAPI(lpMimeStm->CopyTo(lpFileStm, ULARGE_MAX, nullptr, nullptr));
									if (SUCCEEDED(hRes))
									{
										hRes = EC_MAPI(lpFileStm->Commit(STGC_DEFAULT));
									}
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
			CCSFLAGS convertFlags,
			bool bApply,
			HCHARSET hCharSet,
			CSETAPPLYTYPE cSetApplyType,
			_In_opt_ LPADRBOOK lpAdrBook,
			bool bUnicode)
		{
			if (!lpszEMLFile || !lpszMSGFile) return MAPI_E_INVALID_PARAMETER;

			LPSTORAGE pStorage = nullptr;
			LPMESSAGE pMessage = nullptr;

			auto hRes = EC_H(file::CreateNewMSG(lpszMSGFile, bUnicode, &pMessage, &pStorage));
			if (SUCCEEDED(hRes) && pMessage && pStorage)
			{
				hRes = EC_H(ImportEMLToIMessage(
					lpszEMLFile, pMessage, convertFlags, bApply, hCharSet, cSetApplyType, lpAdrBook));
				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI(pStorage->Commit(STGC_DEFAULT));
				}
			}

			if (pMessage) pMessage->Release();
			if (pStorage) pStorage->Release();

			return hRes;
		}

		_Check_return_ HRESULT ConvertMSGToEML(
			_In_z_ LPCWSTR lpszMSGFile,
			_In_z_ LPCWSTR lpszEMLFile,
			CCSFLAGS convertFlags,
			ENCODINGTYPE et,
			MIMESAVETYPE mst,
			ULONG ulWrapLines,
			_In_opt_ LPADRBOOK lpAdrBook)
		{
			if (!lpszEMLFile || !lpszMSGFile) return MAPI_E_INVALID_PARAMETER;

			auto hRes = S_OK;
			auto pMessage = file::LoadMSGToMessage(lpszMSGFile);
			if (pMessage)
			{
				hRes = EC_H(ExportIMessageToEML(pMessage, lpszEMLFile, convertFlags, et, mst, ulWrapLines, lpAdrBook));
				pMessage->Release();
			}

			return hRes;
		}
	} // namespace mapimime
} // namespace mapi