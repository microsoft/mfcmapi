// Collection of useful MAPI functions

#include <StdAfx.h>
#include <UI/MAPIFormFunctions.h>
#include <UI/MyMAPIFormViewer.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

namespace mapi::mapiui
{
	// This function creates a new message of class szMessageClass, based in m_lpContainer
	// The function will also take care of launching the form

	// This function can be used to create a new message using any form.
	// Outlook's default IPM.Note and IPM.Post can be created in any folder, so these don't pose a problem.
	// Appointment, Contact, StickyNote, and Task can only be created in those folders
	// Attempting to create one of those in the Inbox will result in an
	// 'Internal Application Error' when you save.
	_Check_return_ HRESULT CreateAndDisplayNewMailInFolder(
		_In_ HWND hwndParent,
		_In_ LPMDB lpMDB,
		_In_ LPMAPISESSION lpMAPISession,
		_In_ controls::sortlistctrl::CContentsTableListCtrl* lpContentsTableListCtrl,
		int iItem,
		_In_ const std::wstring& szMessageClass,
		_In_ LPMAPIFOLDER lpFolder)
	{
		if (!lpFolder || !lpMAPISession) return MAPI_E_INVALID_PARAMETER;

		LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
		auto hRes = EC_H_MSG(IDS_NOFORMMANAGER, MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));
		if (!lpMAPIFormMgr) return hRes;

		LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
		hRes = EC_H_MSG(
			IDS_NOCLASSHANDLER,
			lpMAPIFormMgr->ResolveMessageClass(
				strings::wstringTostring(szMessageClass).c_str(), // class
				0, // flags
				lpFolder, // folder to resolve to
				&lpMAPIFormInfo));
		if (lpMAPIFormInfo)
		{
			LPPERSISTMESSAGE lpPersistMessage = nullptr;
			hRes = EC_MAPI(lpMAPIFormMgr->CreateForm(
				reinterpret_cast<ULONG_PTR>(hwndParent), // parent window
				MAPI_DIALOG, // display status window
				lpMAPIFormInfo, // form info
				IID_IPersistMessage, // riid to open
				reinterpret_cast<LPVOID*>(&lpPersistMessage))); // form to open into
			if (lpPersistMessage)
			{
				LPMESSAGE lpMessage = nullptr;
				// Get a message
				hRes = EC_MAPI(lpFolder->CreateMessage(
					nullptr, // default interface
					0, // flags
					&lpMessage));
				if (lpMessage)
				{
					auto lpMAPIFormViewer = new (std::nothrow) CMyMAPIFormViewer(
						hwndParent, lpMDB, lpMAPISession, lpFolder, lpMessage, lpContentsTableListCtrl, iItem);

					if (lpMAPIFormViewer)
					{
						// put everything together with the default info
						hRes = EC_MAPI(
							lpPersistMessage->InitNew(static_cast<LPMAPIMESSAGESITE>(lpMAPIFormViewer), lpMessage));

						auto lpForm = mapi::safe_cast<LPMAPIFORM>(lpPersistMessage);
						if (lpForm)
						{
							hRes = EC_MAPI(lpForm->SetViewContext(static_cast<LPMAPIVIEWCONTEXT>(lpMAPIFormViewer)));

							if (SUCCEEDED(hRes))
							{
								hRes = EC_MAPI(lpMAPIFormViewer->CallDoVerb(
									lpForm,
									EXCHIVERB_OPEN,
									nullptr)); // Not passing a RECT here so we'll try to use the default for the form
							}

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
	}

	_Check_return_ HRESULT OpenMessageNonModal(
		_In_ HWND hwndParent,
		_In_ LPMDB lpMDB,
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMAPIFOLDER lpSourceFolder,
		_In_ controls::sortlistctrl::CContentsTableListCtrl* lpContentsTableListCtrl,
		int iItem,
		_In_ LPMESSAGE lpMessage,
		LONG lVerb,
		_In_opt_ LPCRECT lpRect)
	{
		ULONG cValuesShow = 0;
		LPSPropValue lpspvaShow = nullptr;
		ULONG ulMessageStatus = NULL;
		LPMAPIVIEWCONTEXT lpViewContextTemp = nullptr;

		enum
		{
			FLAGS,
			CLASS,
			EID,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptaShowForm) = {NUM_COLS,
																   {PR_MESSAGE_FLAGS, PR_MESSAGE_CLASS_A, PR_ENTRYID}};

		if (!lpMessage || !lpMAPISession || !lpSourceFolder) return MAPI_E_INVALID_PARAMETER;

		// Get required properties from the message
		auto hRes = EC_H_GETPROPS(lpMessage->GetProps(
			LPSPropTagArray(&sptaShowForm), // property tag array
			fMapiUnicode, // flags
			&cValuesShow, // Count of values returned
			&lpspvaShow)); // Values returned

		if (lpspvaShow)
		{
			const auto bin = mapi::getBin(lpspvaShow[EID]);
			hRes = EC_MAPI(
				lpSourceFolder->GetMessageStatus(bin.cb, reinterpret_cast<LPENTRYID>(bin.lpb), 0, &ulMessageStatus));

			auto lpMAPIFormViewer = new (std::nothrow) CMyMAPIFormViewer(
				hwndParent, lpMDB, lpMAPISession, lpSourceFolder, lpMessage, lpContentsTableListCtrl, iItem);

			if (lpMAPIFormViewer)
			{
				LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
				LPMAPIFORM lpForm = nullptr;

				hRes = EC_MAPI(lpMAPIFormViewer->GetFormManager(&lpMAPIFormMgr));

				if (lpMAPIFormMgr)
				{
					output::DebugPrint(
						output::dbgLevel::FormViewer,
						L"Calling LoadForm: szMessageClass = %hs, ulMessageStatus = 0x%X, ulMessageFlags = 0x%X\n",
						lpspvaShow[CLASS].Value.lpszA,
						ulMessageStatus,
						lpspvaShow[FLAGS].Value.ul);
					hRes = EC_MAPI(lpMAPIFormMgr->LoadForm(
						reinterpret_cast<ULONG_PTR>(hwndParent),
						0, // flags
						lpspvaShow[CLASS].Value.lpszA,
						ulMessageStatus,
						lpspvaShow[FLAGS].Value.ul,
						lpSourceFolder,
						lpMAPIFormViewer,
						lpMessage,
						lpMAPIFormViewer,
						IID_IMAPIForm, // riid
						reinterpret_cast<LPVOID*>(&lpForm)));
					lpMAPIFormMgr->Release();
					lpMAPIFormMgr = nullptr;
				}

				if (lpForm)
				{
					hRes = EC_MAPI(lpMAPIFormViewer->CallDoVerb(lpForm, lVerb, lpRect));

					if (SUCCEEDED(hRes))
					{
						// Fix for unknown typed freedocs.
						hRes = WC_MAPI(lpForm->GetViewContext(&lpViewContextTemp));
					}

					if (SUCCEEDED(hRes))
					{
						if (lpViewContextTemp)
						{
							// If we got a pointer back, we'll just release it and continue.
							lpViewContextTemp->Release();
						}
						else
						{
							// If the pointer came back NULL, then we need to call ShutdownForm but don't release.
							hRes = WC_MAPI(lpForm->ShutdownForm(SAVEOPTS_NOSAVE));
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
	}

	_Check_return_ HRESULT OpenMessageModal(
		_In_ LPMAPIFOLDER lpParentFolder,
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMDB lpMDB,
		_In_ LPMESSAGE lpMessage)
	{
		ULONG cValuesShow = 0;
		LPSPropValue lpspvaShow = nullptr;
		ULONG_PTR Token = NULL;
		ULONG ulMessageStatus = NULL;

		enum
		{
			FLAGS,
			CLASS,
			ACCESS,
			EID,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptaShowForm) = {
			NUM_COLS, {PR_MESSAGE_FLAGS, PR_MESSAGE_CLASS_A, PR_ACCESS, PR_ENTRYID}};

		if (!lpMessage || !lpParentFolder || !lpMAPISession || !lpMDB) return MAPI_E_INVALID_PARAMETER;

		// Get required properties from the message
		auto hRes = EC_H_GETPROPS(lpMessage->GetProps(
			LPSPropTagArray(&sptaShowForm), // property tag array
			fMapiUnicode, // flags
			&cValuesShow, // Count of values returned
			&lpspvaShow)); // Values returned

		if (lpspvaShow)
		{
			const auto bin = mapi::getBin(lpspvaShow[EID]);
			hRes = EC_MAPI(
				lpParentFolder->GetMessageStatus(bin.cb, reinterpret_cast<LPENTRYID>(bin.lpb), 0, &ulMessageStatus));

			if (SUCCEEDED(hRes))
			{
				// set up the 'display message' form
				hRes = EC_MAPI(lpMAPISession->PrepareForm(
					nullptr, // default interface
					lpMessage, // message to open
					&Token)); // basically, the pointer to the form
			}

			if (SUCCEEDED(hRes))
			{
				hRes = EC_H_CANCEL(lpMAPISession->ShowForm(
					NULL,
					lpMDB, // message store
					lpParentFolder, // parent folder
					nullptr, // default interface
					Token, // token?
					nullptr, // reserved
					MAPI_POST_MESSAGE, // flags
					ulMessageStatus, // message status
					lpspvaShow[FLAGS].Value.ul, // message flags
					lpspvaShow[ACCESS].Value.ul, // access
					lpspvaShow[CLASS].Value.lpszA)); // message class
			}
		}

		MAPIFreeBuffer(lpspvaShow);
		return hRes;
	}
} // namespace mapi::mapiui