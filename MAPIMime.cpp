#include "stdafx.h"
#include "MAPIMime.h"
#include "File.h"
#include "MAPIFunctions.h"
#include "Guids.h"
#include "Editor.h"
#include "ImportProcs.h"

HRESULT ImportEMLToIMessage(
							LPCTSTR lpszEMLFile,
							LPMESSAGE lpMsg,
							ULONG ulConvertFlags,
							BOOL bApply,
							HCHARSET hCharSet,
							CSETAPPLYTYPE cSetApplyType,
							LPADRBOOK lpAdrBook)
{
	if (!lpszEMLFile || !lpMsg) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	LPCONVERTERSESSION lpConverter = NULL;

	EC_H_MSG(CoCreateInstance(CLSID_IConverterSession,
						NULL,
						CLSCTX_INPROC_SERVER,
						IID_IConverterSession,
						(LPVOID*)&lpConverter),IDS_NOCONVERTERSESSION);
	if (SUCCEEDED(hRes) && lpConverter)
	{
		LPSTREAM lpEMLStm = NULL;

		EC_H(MyOpenStreamOnFile(MAPIAllocateBuffer,
			MAPIFreeBuffer,
			STGM_READ,
			(LPTSTR)lpszEMLFile,
			NULL,
			&lpEMLStm));
		if (SUCCEEDED(hRes) && lpEMLStm)
		{
			if (lpAdrBook)
			{
				EC_H(lpConverter->SetAdrBook(lpAdrBook));
			}
			if (SUCCEEDED(hRes) && bApply)
			{
				EC_H(lpConverter->SetCharset(bApply,hCharSet,cSetApplyType));
			}
			if (SUCCEEDED(hRes))
			{
				// We'll make the user ensure CCSF_SMTP is passed
				EC_H(lpConverter->MIMEToMAPI(lpEMLStm,
					lpMsg,
					NULL, // Must be NULL
					ulConvertFlags));
				if (SUCCEEDED(hRes))
				{
					EC_H(lpMsg->SaveChanges(NULL));
				}
			}
		}

		if (lpEMLStm) lpEMLStm->Release();
	}
	// If we failed to get the converter, just return OK so we don't keep getting errors.
	if (REGDB_E_CLASSNOTREG == hRes) hRes = S_OK;

	if (lpConverter) lpConverter->Release();

	return hRes;
}

HRESULT ExportIMessageToEML(LPMESSAGE lpMsg, LPCTSTR lpszEMLFile, ULONG ulConvertFlags,
							ENCODINGTYPE et, MIMESAVETYPE mst, ULONG ulWrapLines,LPADRBOOK lpAdrBook)
{
	if (!lpszEMLFile || !lpMsg) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	LPCONVERTERSESSION lpConverter = NULL;

	EC_H_MSG(CoCreateInstance(CLSID_IConverterSession,
						NULL,
						CLSCTX_INPROC_SERVER,
						IID_IConverterSession,
						(LPVOID*)&lpConverter),IDS_NOCONVERTERSESSION);
	if (SUCCEEDED(hRes) && lpConverter)
	{
		if (lpAdrBook)
		{
			EC_H(lpConverter->SetAdrBook(lpAdrBook));
		}
		if (SUCCEEDED(hRes) && et != IET_UNKNOWN)
		{
			EC_H(lpConverter->SetEncoding(et));
		}
		if (SUCCEEDED(hRes) && mst != USE_DEFAULT_SAVETYPE)
		{
			EC_H(lpConverter->SetSaveFormat(mst));
		}
		if (SUCCEEDED(hRes) && ulWrapLines != USE_DEFAULT_WRAPPING)
		{
			EC_H(lpConverter->SetTextWrapping(ulWrapLines != 0 ? true : false, ulWrapLines));
		}

		if (SUCCEEDED(hRes))
		{
			LPSTREAM lpMimeStm = NULL;

			EC_H(CreateStreamOnHGlobal(NULL, TRUE, &lpMimeStm));
			if (SUCCEEDED(hRes) && lpMimeStm)
			{
				// Per the docs for MAPIToMIMEStm, CCSF_SMTP must always be set
				// But we're gonna make the user ensure that, so we don't or it in here
				EC_H(lpConverter->MAPIToMIMEStm(lpMsg, lpMimeStm, ulConvertFlags));
				if (SUCCEEDED(hRes))
				{
					LPSTREAM lpFileStm = NULL;

					EC_H(MyOpenStreamOnFile(MAPIAllocateBuffer,
											MAPIFreeBuffer,
											STGM_CREATE | STGM_READWRITE,
											(LPTSTR)lpszEMLFile,
											NULL,
											&lpFileStm));
					if (SUCCEEDED(hRes) && lpFileStm)
					{
						LARGE_INTEGER dwBegin = {0};
						EC_H(lpMimeStm->Seek(dwBegin, STREAM_SEEK_SET, NULL));
						if (SUCCEEDED(hRes))
						{
							EC_H(lpMimeStm->CopyTo(lpFileStm, ULARGE_MAX, NULL, NULL));
							if (SUCCEEDED(hRes))
							{
								EC_H(lpFileStm->Commit(STGC_DEFAULT));
							}
						}
					}

					if (lpFileStm) lpFileStm->Release();
				}
			}

			if (lpMimeStm) lpMimeStm->Release();
		}
	}
	// If we failed to get the converter, just return OK so we don't keep getting errors.
	if (REGDB_E_CLASSNOTREG == hRes) hRes = S_OK;

	if (lpConverter) lpConverter->Release();

	return hRes;
}

HRESULT ConvertEMLToMSG(
						LPCTSTR lpszEMLFile,
						LPCTSTR lpszMSGFile,
						ULONG ulConvertFlags,
						BOOL bApply,
						HCHARSET hCharSet,
						CSETAPPLYTYPE cSetApplyType,
						LPADRBOOK lpAdrBook,
						BOOL bUnicode)
{
	if (!lpszEMLFile || !lpszMSGFile) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	LPSTORAGE pStorage = NULL;
	LPMESSAGE pMessage = NULL;

	EC_H(CreateNewMSG(lpszMSGFile, bUnicode, &pMessage, &pStorage));
	if (SUCCEEDED(hRes) && pMessage && pStorage)
	{
		EC_H(ImportEMLToIMessage(
			lpszEMLFile,
			pMessage,
			ulConvertFlags,
			bApply,
			hCharSet,
			cSetApplyType,
			lpAdrBook));
		if (SUCCEEDED(hRes))
		{
			EC_H(pStorage->Commit(STGC_DEFAULT));
		}
	}

	if (pMessage) pMessage->Release();
	if (pStorage) pStorage->Release();

	return hRes;
}

HRESULT ConvertMSGToEML(LPCTSTR lpszMSGFile,
						LPCTSTR lpszEMLFile,
						ULONG ulConvertFlags,
						ENCODINGTYPE et,
						MIMESAVETYPE mst,
						ULONG ulWrapLines,
						LPADRBOOK lpAdrBook)
{
	if (!lpszEMLFile || !lpszMSGFile) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	LPMESSAGE pMessage = NULL;

	EC_H(LoadMSGToMessage(lpszMSGFile, &pMessage));
	if (SUCCEEDED(hRes) && pMessage)
	{
		EC_H(ExportIMessageToEML(
			pMessage,
			lpszEMLFile,
			ulConvertFlags,
			et,
			mst,
			ulWrapLines,
			lpAdrBook));
	}

	if (pMessage) pMessage->Release();

	return hRes;
}

HRESULT GetConversionToEMLOptions(CWnd* pParentWnd,
								  ULONG* lpulConvertFlags,
								  ENCODINGTYPE* lpet,
								  MIMESAVETYPE* lpmst,
								  ULONG* lpulWrapLines,
								  BOOL* pDoAdrBook)
{
	if (!lpulConvertFlags || !lpet || !lpmst || !lpulWrapLines || !pDoAdrBook) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;

	CEditor MyData(
		pParentWnd,
		IDS_CONVERTTOEML,
		IDS_CONVERTTOEMLPROMPT,
		8,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_CONVERTFLAGS,NULL,false);
	MyData.SetHex(0,CCSF_SMTP);
	MyData.InitCheck(1,IDS_CONVERTDOENCODINGTYPE,false,false);
	MyData.InitSingleLine(2,IDS_CONVERTENCODINGTYPE,NULL,false);
	MyData.SetHex(2,IET_7BIT);
	MyData.InitCheck(3,IDS_CONVERTDOMIMESAVETYPE,false,false);
	MyData.InitSingleLine(4,IDS_CONVERTMIMESAVETYPE,NULL,false);
	MyData.SetHex(4,SAVE_RFC822);
	MyData.InitCheck(5,IDS_CONVERTDOWRAPLINES,false,false);
	MyData.InitSingleLine(6,IDS_CONVERTWRAPLINECOUNT,NULL,false);
	MyData.SetDecimal(6,74);
	MyData.InitCheck(7,IDS_CONVERTDOADRBOOK,false,false);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		*lpulConvertFlags = MyData.GetHex(0);
		*lpulWrapLines =  MyData.GetCheck(1)?(ENCODINGTYPE)MyData.GetDecimal(2):IET_UNKNOWN;
		*lpmst =  MyData.GetCheck(3)?(MIMESAVETYPE)MyData.GetHex(4):USE_DEFAULT_SAVETYPE;
		*lpulWrapLines = MyData.GetCheck(5)?MyData.GetDecimal(6):USE_DEFAULT_WRAPPING;
		*pDoAdrBook = MyData.GetCheck(7);
	}
	return hRes;
} // GetConversionToEMLOptions

HRESULT GetConversionFromEMLOptions(CWnd* pParentWnd,
									ULONG* lpulConvertFlags,
									BOOL* pDoAdrBook,
									BOOL* pDoApply,
									HCHARSET* phCharSet,
									CSETAPPLYTYPE* pcSetApplyType,
									BOOL* pbUnicode)
{
	if (!lpulConvertFlags || !pDoAdrBook || !pDoApply || !phCharSet || !pcSetApplyType) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;

	CEditor MyData(
		pParentWnd,
		IDS_CONVERTFROMEML,
		IDS_CONVERTFROMEMLPROMPT,
		pbUnicode?7:6,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_CONVERTFLAGS,NULL,false);
	MyData.SetHex(0,CCSF_SMTP);
	MyData.InitCheck(1,IDS_CONVERTCODEPAGE,false,false);
	MyData.InitSingleLine(2,IDS_CONVERTCODEPAGE,NULL,false);
	MyData.SetDecimal(2,CP_USASCII);
	MyData.InitSingleLine(3,IDS_CONVERTCHARSETTYPE,NULL,false);
	MyData.SetDecimal(3,CHARSET_BODY);
	MyData.InitSingleLine(4,IDS_CONVERTCHARSETAPPLYTYPE,NULL,false);
	MyData.SetDecimal(4,CSET_APPLY_UNTAGGED);
	MyData.InitCheck(5,IDS_CONVERTDOADRBOOK,false,false);
	if (pbUnicode)
	{
		MyData.InitCheck(6,IDS_SAVEUNICODE,false,false);
	}

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		*lpulConvertFlags = MyData.GetHex(0);
		if (MyData.GetCheck(1))
		{
			if (!pfnMimeOleGetCodePageCharset) return MAPI_E_INVALID_PARAMETER;
			if (SUCCEEDED(hRes)) *pDoApply = true;
			*pcSetApplyType = (CSETAPPLYTYPE) MyData.GetDecimal(4);
			if (*pcSetApplyType > CSET_APPLY_TAG_ALL) return MAPI_E_INVALID_PARAMETER;
			ULONG ulCodePage = MyData.GetDecimal(2);
			CHARSETTYPE cCharSetType = (CHARSETTYPE) MyData.GetDecimal(3);
			if (cCharSetType > CHARSET_WEB) return MAPI_E_INVALID_PARAMETER;
			EC_H(pfnMimeOleGetCodePageCharset(ulCodePage,cCharSetType,phCharSet));
		}
		*pDoAdrBook = MyData.GetCheck(5);
		if (pbUnicode)
		{
			*pbUnicode = MyData.GetCheck(6);
		}
	}
	return hRes;
}