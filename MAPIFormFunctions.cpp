// MAPIFormfunctions.cpp : Collection of useful MAPI functions

#include "stdafx.h"
#include "MAPIFormFunctions.h"
#include "MyMAPIFormViewer.h"
#include "MAPIFunctions.h"

// This function creates a new message of class szMessageClass, based in m_lpContainer
// The function will also take care of launching the form

// This function can be used to create a new message using any form.
// Outlook's default IPM.Note and IPM.Post can be created in any folder, so these don't pose a problem.
// Appointment, Contact, StickyNote, and Task can only be created in those folders
// Attempting to create one of those in the Inbox will result in an
// 'Internal Application Error' when you save.

HRESULT CreateAndDisplayNewMailInFolder(
							  LPMDB lpMDB,
							  LPMAPISESSION lpMAPISession,
							  CContentsTableListCtrl *lpContentsTableListCtrl,
							  int iItem,
							  LPCSTR szMessageClass,
							  LPMAPIFOLDER lpFolder)
{
	HRESULT				hRes = S_OK;
	LPMAPIFORMMGR		lpMAPIFormMgr = NULL;

	if (!lpFolder || !lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	EC_H_MSG(MAPIOpenFormMgr(lpMAPISession,&lpMAPIFormMgr),IDS_NOFORMMANAGER);

	if (!lpMAPIFormMgr) return hRes;

	LPMAPIFORMINFO		lpMAPIFormInfo = NULL;
	LPPERSISTMESSAGE	lpPersistMessage = NULL;

	EC_H_MSG(lpMAPIFormMgr->ResolveMessageClass(
		szMessageClass, // class
		NULL, // flags
		lpFolder, // folder to resolve to
		&lpMAPIFormInfo),
		IDS_NOCLASSHANDLER);
	if (lpMAPIFormInfo)
	{
		EC_H(lpMAPIFormMgr->CreateForm(
			NULL, // parent window
			MAPI_DIALOG, // display status window
			lpMAPIFormInfo, // form info
			IID_IPersistMessage, // riid to open
			(LPVOID *) &lpPersistMessage)); // form to open into

		if (lpPersistMessage)
		{
			LPMESSAGE lpMessage = NULL;
			// Get a message
			EC_H(lpFolder->CreateMessage(
				NULL, // default interface
				0, // flags
				&lpMessage));
			if (lpMessage)
			{
				CMyMAPIFormViewer*	lpMAPIFormViewer = NULL;
				lpMAPIFormViewer = new CMyMAPIFormViewer(
					NULL,
					lpMDB,
					lpMAPISession,
					lpFolder,
					lpMessage,
					lpContentsTableListCtrl,
					iItem);

				if (lpMAPIFormViewer)
				{
					// put everything together with the default info
					EC_H(lpPersistMessage->InitNew(
						(LPMAPIMESSAGESITE) lpMAPIFormViewer,
						lpMessage));

					LPMAPIFORM lpForm = NULL;
					EC_H(lpPersistMessage->QueryInterface(IID_IMAPIForm,(LPVOID*) &lpForm));

					if (lpForm)
					{
						EC_H(lpForm->SetViewContext(
							(LPMAPIVIEWCONTEXT) lpMAPIFormViewer));

						EC_H(lpMAPIFormViewer->CallDoVerb(
							lpForm,
							EXCHIVERB_OPEN,
							NULL)); // Not passing a RECT here so we'll try to use the default for the form
						lpForm->Release();
					}
					lpMAPIFormViewer->Release();
				}
				lpMessage->Release();
			}
			lpPersistMessage->Release();
		}
		lpMAPIFormInfo->Release();
	}
	lpMAPIFormMgr->Release();
	return hRes;
} // CreateAndDisplayNewMailInFolder

HRESULT OpenMessageNonModal(
							LPMDB lpMDB,
							LPMAPISESSION lpMAPISession,
							LPMAPIFOLDER lpSourceFolder,
							CContentsTableListCtrl *lpContentsTableListCtrl,
							int iItem,
							LPMESSAGE lpMessage,
							LONG lVerb,
							LPCRECT lpRect)
{
	HRESULT					hRes = S_OK;
	ULONG					cValuesShow = 0;
	LPSPropValue			lpspvaShow = NULL;
	ULONG					ulMessageStatus = NULL;
	LPMAPIVIEWCONTEXT		lpViewContextTemp = NULL;

	enum {FLAGS,CLASS,EID,NUM_COLS};
	SizedSPropTagArray(NUM_COLS,sptaShowForm) = { NUM_COLS, {
		PR_MESSAGE_FLAGS,
			PR_MESSAGE_CLASS_A,
			PR_ENTRYID}
	};

	if (!lpMessage || !lpMAPISession || !lpSourceFolder) return MAPI_E_INVALID_PARAMETER;

	// Get required properties from the message
	EC_H_GETPROPS(lpMessage->GetProps(
		(LPSPropTagArray) &sptaShowForm, // property tag array
		fMapiUnicode, // flags
		&cValuesShow, // Count of values returned
		&lpspvaShow)); // Values returned

	if (lpspvaShow)
	{
		EC_H(lpSourceFolder->GetMessageStatus(
			lpspvaShow[EID].Value.bin.cb,
			(LPENTRYID)lpspvaShow[EID].Value.bin.lpb,
			0,
			&ulMessageStatus));

		CMyMAPIFormViewer* lpMAPIFormViewer = NULL;
		lpMAPIFormViewer = new CMyMAPIFormViewer(
			NULL,
			lpMDB,
			lpMAPISession,
			lpSourceFolder,
			lpMessage,
			lpContentsTableListCtrl,
			iItem);

		if (lpMAPIFormViewer)
		{
			LPMAPIFORMMGR lpMAPIFormMgr = NULL;
			LPMAPIFORM lpForm = NULL;

			EC_H(lpMAPIFormViewer->GetFormManager(&lpMAPIFormMgr));

			if (lpMAPIFormMgr)
			{
				DebugPrint(DBGFormViewer,_T("Calling LoadForm: szMessageClass = %hs, ulMessageStatus = 0x%X, ulMessageFlags = 0x%X\n"),
					lpspvaShow[CLASS].Value.lpszA,
					ulMessageStatus,
					lpspvaShow[FLAGS].Value.ul);
				EC_H(lpMAPIFormMgr->LoadForm(
					0, // (ULONG) m_hWnd,
					0, // flags
					lpspvaShow[CLASS].Value.lpszA,
					ulMessageStatus,
					lpspvaShow[FLAGS].Value.ul,
					lpSourceFolder,
					lpMAPIFormViewer,
					lpMessage,
					lpMAPIFormViewer,
					IID_IMAPIForm, // riid
					(LPVOID *) &lpForm));
				lpMAPIFormMgr->Release();
				lpMAPIFormMgr = NULL;
			}

			if (lpForm)
			{
				EC_H(lpMAPIFormViewer->CallDoVerb(
					lpForm,
					lVerb,
					lpRect));
				// Fix for unknown typed freedocs.
				WC_H(lpForm->GetViewContext(&lpViewContextTemp));
				if (SUCCEEDED(hRes)){
					if (lpViewContextTemp){
						// If we got a pointer back, we'll just release it and continue.
						lpViewContextTemp->Release();
					}
					else{
						// If the pointer came back NULL, then we need to call ShutdownForm but don't release.
						WC_H(lpForm->ShutdownForm(SAVEOPTS_NOSAVE));
					}
				}
				else
				{
					// Not getting a view context isn't a bad thing
					hRes = S_OK;
				}

				lpForm->Release();
			}
			lpMAPIFormViewer->Release();
		}

		MAPIFreeBuffer(lpspvaShow);
	}
	return hRes;
} // OpenMessageNonModal


HRESULT OpenMessageModal(LPMAPIFOLDER lpParentFolder,
						 LPMAPISESSION lpMAPISession,
						 LPMDB lpMDB,
						 LPMESSAGE lpMessage)
{
	HRESULT			hRes = S_OK;
	ULONG			cValuesShow;
	LPSPropValue	lpspvaShow = NULL;
	ULONG_PTR		Token = NULL;
	ULONG			ulMessageStatus = NULL;

	enum {FLAGS,CLASS,ACCESS,EID,NUM_COLS};
	SizedSPropTagArray(NUM_COLS,sptaShowForm) = { NUM_COLS, {
		PR_MESSAGE_FLAGS,
			PR_MESSAGE_CLASS_A,
			PR_ACCESS,
			PR_ENTRYID}
	};

	if (!lpMessage || !lpParentFolder || !lpMAPISession || !lpMDB) return MAPI_E_INVALID_PARAMETER;

	// Get required properties from the message
	EC_H_GETPROPS(lpMessage->GetProps(
		(LPSPropTagArray) &sptaShowForm, // property tag array
		fMapiUnicode, // flags
		&cValuesShow, // Count of values returned
		&lpspvaShow)); // Values returned

	if (lpspvaShow)
	{
		EC_H(lpParentFolder->GetMessageStatus(
			lpspvaShow[EID].Value.bin.cb,
			(LPENTRYID)lpspvaShow[EID].Value.bin.lpb,
			0,
			&ulMessageStatus));

		// set up the 'display message' form
		EC_H(lpMAPISession->PrepareForm(
			NULL, // default interface
			lpMessage, // message to open
			&Token)); // basically, the pointer to the form

		EC_H_CANCEL(lpMAPISession->ShowForm(
			NULL,
			lpMDB, // message store
			lpParentFolder, // parent folder
			NULL, // default interface
			Token, // token?
			NULL, // reserved
			MAPI_POST_MESSAGE, // flags
			ulMessageStatus, // message status
			lpspvaShow[FLAGS].Value.ul, // message flags
			lpspvaShow[ACCESS].Value.ul, // access
			lpspvaShow[CLASS].Value.lpszA)); // message class
	}

	MAPIFreeBuffer(lpspvaShow);
	return hRes;
} // OpenMessageModal