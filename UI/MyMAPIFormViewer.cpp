// Implementation of the CMyMAPIFormViewer class.

#include <StdAfx.h>
#include <UI/MyMAPIFormViewer.h>
#include <UI/MAPIFormFunctions.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <core/interpret/guid.h>
#include <UI/MyWinApp.h>
#include <core/sortlistdata/contentsData.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

extern ui::CMyWinApp theApp;

namespace mapi
{
	namespace mapiui
	{
		static std::wstring CLASS = L"CMyMAPIFormViewer";

		CMyMAPIFormViewer::CMyMAPIFormViewer(
			_In_ HWND hwndParent,
			_In_ LPMDB lpMDB,
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMAPIFOLDER lpFolder,
			_In_ LPMESSAGE lpMessage,
			_In_opt_ controls::sortlistctrl::CContentsTableListCtrl* lpContentsTableListCtrl,
			int iItem)
		{
			TRACE_CONSTRUCTOR(CLASS);

			m_cRef = 1;
			m_iItem = iItem;

			m_lpMDB = lpMDB;
			m_lpMAPISession = lpMAPISession;
			m_lpFolder = lpFolder;
			m_lpMessage = lpMessage;
			m_lpContentsTableListCtrl = lpContentsTableListCtrl;

			if (m_lpMDB) m_lpMDB->AddRef();
			if (m_lpMAPISession) m_lpMAPISession->AddRef();
			if (m_lpFolder) m_lpFolder->AddRef();
			if (m_lpMessage) m_lpMessage->AddRef();
			if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->AddRef();

			m_lpPersistMessage = nullptr;
			m_lpMapiFormAdviseSink = nullptr;

			m_hwndParent = hwndParent;
			if (!m_hwndParent)
			{
				m_hwndParent = GetActiveWindow();
			}

			if (!m_hwndParent && theApp.m_pMainWnd)
			{
				m_hwndParent = theApp.m_pMainWnd->m_hWnd;
			}

			if (!m_hwndParent)
			{
				error::ErrDialog(__FILE__, __LINE__, IDS_EDFORMVIEWERNULLPARENT);
				output::DebugPrint(output::DBGFormViewer, L"Form Viewer created with a NULL parent!\n");
			}
		}

		CMyMAPIFormViewer::~CMyMAPIFormViewer()
		{
			TRACE_DESTRUCTOR(CLASS);
			ReleaseObjects();
		}

		void CMyMAPIFormViewer::ReleaseObjects()
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"ReleaseObjects", L"\n");
			ShutdownPersist();
			if (m_lpMessage) m_lpMessage->Release();
			if (m_lpFolder) m_lpFolder->Release();
			if (m_lpMDB) m_lpMDB->Release();
			if (m_lpMAPISession) m_lpMAPISession->Release();
			if (m_lpMapiFormAdviseSink) m_lpMapiFormAdviseSink->Release();
			if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->Release(); // this must be last!!!!!
			m_lpMessage = nullptr;
			m_lpFolder = nullptr;
			m_lpMDB = nullptr;
			m_lpMAPISession = nullptr;
			m_lpMapiFormAdviseSink = nullptr;
			m_lpContentsTableListCtrl = nullptr;
		}

		STDMETHODIMP CMyMAPIFormViewer::QueryInterface(REFIID riid, LPVOID* ppvObj)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"QueryInterface", L"\n");
			auto szGuid = guid::GUIDToStringAndName(const_cast<LPGUID>(&riid));
			output::DebugPrint(output::DBGFormViewer, L"GUID Requested: %ws\n", szGuid.c_str());

			*ppvObj = nullptr;
			if (riid == IID_IMAPIMessageSite)
			{
				output::DebugPrint(output::DBGFormViewer, L"Requested IID_IMAPIMessageSite\n");
				*ppvObj = static_cast<IMAPIMessageSite*>(this);
				AddRef();
				return S_OK;
			}

			if (riid == IID_IMAPIViewContext)
			{
				output::DebugPrint(output::DBGFormViewer, L"Requested IID_IMAPIViewContext\n");
				*ppvObj = static_cast<IMAPIViewContext*>(this);
				AddRef();
				return S_OK;
			}

			if (riid == IID_IMAPIViewAdviseSink)
			{
				output::DebugPrint(output::DBGFormViewer, L"Requested IID_IMAPIViewAdviseSink\n");
				*ppvObj = static_cast<IMAPIViewAdviseSink*>(this);
				AddRef();
				return S_OK;
			}

			if (riid == IID_IUnknown)
			{
				output::DebugPrint(output::DBGFormViewer, L"Requested IID_IUnknown\n");
				*ppvObj = static_cast<LPUNKNOWN>(static_cast<IMAPIMessageSite*>(this));
				AddRef();
				return S_OK;
			}

			output::DebugPrint(output::DBGFormViewer, L"Unknown interface requested\n");
			return E_NOINTERFACE;
		}

		STDMETHODIMP_(ULONG) CMyMAPIFormViewer::AddRef()
		{
			const auto lCount = InterlockedIncrement(&m_cRef);
			TRACE_ADDREF(CLASS, lCount);
			return lCount;
		}

		STDMETHODIMP_(ULONG) CMyMAPIFormViewer::Release()
		{
			const auto lCount = InterlockedDecrement(&m_cRef);
			TRACE_RELEASE(CLASS, lCount);
			if (!lCount) delete this;
			return lCount;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetSession(LPMAPISESSION* ppSession)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetSession", L"\n");
			if (ppSession) *ppSession = m_lpMAPISession;
			if (ppSession && *ppSession)
			{
				m_lpMAPISession->AddRef();
				return S_OK;
			}

			return S_FALSE;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetStore(LPMDB* ppStore)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetStore", L"\n");
			if (ppStore) *ppStore = m_lpMDB;
			if (ppStore && *ppStore)
			{
				m_lpMDB->AddRef();
				return S_OK;
			}

			return S_FALSE;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetFolder(LPMAPIFOLDER* ppFolder)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetFolder", L"\n");
			if (ppFolder) *ppFolder = m_lpFolder;
			if (ppFolder && *ppFolder)
			{
				m_lpFolder->AddRef();
				return S_OK;
			}

			return S_FALSE;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetMessage(LPMESSAGE* ppmsg)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetMessage", L"\n");
			if (ppmsg) *ppmsg = m_lpMessage;
			if (ppmsg && *ppmsg)
			{
				m_lpMessage->AddRef();
				return S_OK;
			}

			return S_FALSE;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetFormManager(LPMAPIFORMMGR* ppFormMgr)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetFormManager", L"\n");
			const auto hRes = EC_MAPI(MAPIOpenFormMgr(m_lpMAPISession, ppFormMgr));
			return hRes;
		}

		STDMETHODIMP CMyMAPIFormViewer::NewMessage(
			ULONG fComposeInFolder,
			LPMAPIFOLDER pFolderFocus,
			LPPERSISTMESSAGE pPersistMessage,
			LPMESSAGE* ppMessage,
			LPMAPIMESSAGESITE* ppMessageSite,
			LPMAPIVIEWCONTEXT* ppViewContext)
		{
			output::DebugPrintEx(
				output::DBGFormViewer,
				CLASS,
				L"NewMessage",
				L"fComposeInFolder = 0x%X pFolderFocus = %p, pPersistMessage = %p\n",
				fComposeInFolder,
				pFolderFocus,
				pPersistMessage);
			auto hRes = S_OK;

			*ppMessage = nullptr;
			*ppMessageSite = nullptr;
			if (ppViewContext) *ppViewContext = nullptr;

			if (!static_cast<bool>(fComposeInFolder) || !pFolderFocus)
			{
				pFolderFocus = m_lpFolder;
			}

			if (pFolderFocus)
			{
				hRes = EC_MAPI(pFolderFocus->CreateMessage(
					nullptr, // IID
					NULL, // flags
					ppMessage));

				if (*ppMessage) // not going to release this because we're returning it
				{
					// don't free since we're passing it back
					auto lpMAPIFormViewer = new (std::nothrow) CMyMAPIFormViewer(
						m_hwndParent,
						m_lpMDB,
						m_lpMAPISession,
						pFolderFocus,
						*ppMessage,
						nullptr, // m_lpContentsTableListCtrl, // don't need this on a new message
						-1);
					if (lpMAPIFormViewer) // not going to release this because we're returning it in ppMessageSite
					{
						hRes = EC_H(lpMAPIFormViewer->SetPersist(nullptr, pPersistMessage));
						*ppMessageSite = static_cast<LPMAPIMESSAGESITE>(lpMAPIFormViewer);
					}
				}
			}

			return hRes;
		}

		STDMETHODIMP CMyMAPIFormViewer::CopyMessage(LPMAPIFOLDER /*pFolderDestination*/)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"CopyMessage", L"\n");
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::MoveMessage(
			LPMAPIFOLDER /*pFolderDestination*/,
			LPMAPIVIEWCONTEXT /*pViewContext*/,
			LPCRECT /*prcPosRect*/)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"MoveMessage", L"\n");
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::DeleteMessage(LPMAPIVIEWCONTEXT /*pViewContext*/, LPCRECT /*prcPosRect*/)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"DeleteMessage", L"\n");
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::SaveMessage()
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"SaveMessage", L"\n");

			if (!m_lpPersistMessage || !m_lpMessage) return MAPI_E_INVALID_PARAMETER;

			auto hRes = EC_MAPI(m_lpPersistMessage->Save(
				nullptr, // m_lpMessage,
				true));
			if (FAILED(hRes))
			{
				LPMAPIERROR lpErr = nullptr;
				hRes = WC_MAPI(m_lpPersistMessage->GetLastError(hRes, fMapiUnicode, &lpErr));
				if (lpErr)
				{
					EC_MAPIERR(fMapiUnicode, lpErr);
					MAPIFreeBuffer(lpErr);
				}
				else
					CHECKHRES(hRes);
			}
			else
			{
				hRes = EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI(m_lpPersistMessage->SaveCompleted(nullptr));
				}
			}

			return hRes;
		}

		STDMETHODIMP CMyMAPIFormViewer::SubmitMessage(ULONG ulFlags)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"SubmitMessage", L"ulFlags = 0x%08X\n", ulFlags);
			if (!m_lpPersistMessage || !m_lpMessage) return MAPI_E_INVALID_PARAMETER;

			auto hRes = EC_MAPI(m_lpPersistMessage->Save(m_lpMessage, true));
			if (FAILED(hRes))
			{
				LPMAPIERROR lpErr = nullptr;
				hRes = WC_MAPI(m_lpPersistMessage->GetLastError(hRes, fMapiUnicode, &lpErr));
				if (lpErr)
				{
					EC_MAPIERR(fMapiUnicode, lpErr);
					MAPIFreeBuffer(lpErr);
				}
				else
					CHECKHRES(hRes);
			}
			else
			{
				hRes = EC_MAPI(m_lpPersistMessage->HandsOffMessage());

				if (SUCCEEDED(hRes))
				{
					EC_MAPI_S(m_lpMessage->SubmitMessage(NULL));
				}
			}

			m_lpMessage->Release();
			m_lpMessage = nullptr;
			return S_OK;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetSiteStatus(LPULONG lpulStatus)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetSiteStatus", L"\n");
			*lpulStatus = VCSTATUS_NEW_MESSAGE;
			if (m_lpPersistMessage)
			{
				*lpulStatus |= VCSTATUS_SAVE | VCSTATUS_SUBMIT | NULL;
			}

			return S_OK;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetLastError(HRESULT hResult, ULONG ulFlags, LPMAPIERROR* /*lppMAPIError*/)
		{
			output::DebugPrintEx(
				output::DBGFormViewer,
				CLASS,
				L"GetLastError",
				L"hResult = 0x%08X, ulFlags = 0x%08X\n",
				hResult,
				ulFlags);
			return MAPI_E_NO_SUPPORT;
		}

		// Assuming we've advised on this form, we need to Unadvise it now, or it will never unload
		STDMETHODIMP CMyMAPIFormViewer::OnShutdown()
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"OnShutdown", L"\n");
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::OnNewMessage()
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"OnNewMessage", L"\n");
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::OnPrint(ULONG dwPageNumber, HRESULT hrStatus)
		{
			output::DebugPrintEx(
				output::DBGFormViewer,
				CLASS,
				L"OnPrint",
				L"Page Number %u\nStatus: 0x%08X\n", // STRING_OK
				dwPageNumber,
				hrStatus);
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::OnSubmitted()
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"OnSubmitted", L"\n");
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::OnSaved()
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"OnSaved", L"\n");
			return MAPI_E_NO_SUPPORT;
		}

		// Only call this when we really plan on shutting down
		void CMyMAPIFormViewer::ShutdownPersist()
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"ShutdownPersist", L"\n");

			if (m_lpPersistMessage)
			{
				m_lpPersistMessage->HandsOffMessage();
				m_lpPersistMessage->Release();
				m_lpPersistMessage = nullptr;
			}
		}

		// Used to set or clear our m_lpPersistMessage
		// Can take either a persist or a form interface, but not both
		// This should only be called with a persist if the verb is EXCHIVERB_OPEN
		_Check_return_ HRESULT
		CMyMAPIFormViewer::SetPersist(_In_opt_ LPMAPIFORM lpForm, _In_opt_ LPPERSISTMESSAGE lpPersist)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"SetPersist", L"\n");
			ShutdownPersist();

			static const SizedSPropTagArray(1, sptaFlags) = {1, {PR_MESSAGE_FLAGS}};
			ULONG cValues = 0L;
			LPSPropValue lpPropArray = nullptr;

			const auto hRes = EC_MAPI(m_lpMessage->GetProps(LPSPropTagArray(&sptaFlags), 0, &cValues, &lpPropArray));
			const auto bComposing = lpPropArray && lpPropArray->Value.l & MSGFLAG_UNSENT;
			MAPIFreeBuffer(lpPropArray);

			if (bComposing && !registry::allowPersistCache)
			{
				// Message is in compose mode - caching the IPersistMessage might hose Outlook
				// However, if we don't cache the IPersistMessage, we can't support SaveMessage or SubmitMessage
				output::DebugPrint(output::DBGFormViewer, L"Not caching the persist\n");
				return S_OK;
			}

			output::DebugPrint(output::DBGFormViewer, L"Caching the persist\n");

			// Doing this part (saving the persist) will leak Winword in compose mode - which is why we make the above check
			// trouble is - compose mode is when we need the persist, for SaveMessage and SubmitMessage
			if (lpPersist)
			{
				m_lpPersistMessage = lpPersist;
				m_lpPersistMessage->AddRef();
			}
			else if (lpForm)
			{
				m_lpPersistMessage = mapi::safe_cast<LPPERSISTMESSAGE>(lpForm);
			}

			return hRes;
		}

		_Check_return_ HRESULT
		CMyMAPIFormViewer::CallDoVerb(_In_ LPMAPIFORM lpMapiForm, LONG lVerb, _In_opt_ LPCRECT lpRect)
		{
			auto hRes = S_OK;
			if (EXCHIVERB_OPEN == lVerb)
			{
				hRes = WC_H(SetPersist(lpMapiForm, nullptr));
			}

			if (SUCCEEDED(hRes))
			{
				hRes = WC_MAPI(lpMapiForm->DoVerb(
					lVerb,
					nullptr, // view context
					reinterpret_cast<ULONG_PTR>(m_hwndParent), // parent window
					lpRect)); // RECT structure with size
				if (hRes != S_OK)
				{
					RECT Rect;

					Rect.left = 0;
					Rect.right = 500;
					Rect.top = 0;
					Rect.bottom = 400;
					hRes = EC_MAPI(lpMapiForm->DoVerb(
						lVerb,
						nullptr, // view context
						reinterpret_cast<ULONG_PTR>(m_hwndParent), // parent window
						&Rect)); // RECT structure with size
				}
			}

			return hRes;
		}

		STDMETHODIMP CMyMAPIFormViewer::SetAdviseSink(LPMAPIFORMADVISESINK pmvns)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"SetAdviseSink", L"pmvns = %p\n", pmvns);
			if (m_lpMapiFormAdviseSink) m_lpMapiFormAdviseSink->Release();
			if (pmvns)
			{
				m_lpMapiFormAdviseSink = pmvns;
				m_lpMapiFormAdviseSink->AddRef();
			}
			else
			{
				m_lpMapiFormAdviseSink = nullptr;
			}

			return S_OK;
		}

		STDMETHODIMP CMyMAPIFormViewer::ActivateNext(ULONG ulDir, LPCRECT lpRect)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"ActivateNext", L"ulDir = 0x%X\n", ulDir);

			enum
			{
				ePR_MESSAGE_FLAGS,
				ePR_MESSAGE_CLASS_A,
				NUM_COLS
			};
			static const SizedSPropTagArray(NUM_COLS, sptaShowForm) = {NUM_COLS,
																	   {PR_MESSAGE_FLAGS, PR_MESSAGE_CLASS_A}};

			auto iNewItem = -1;
			LPMESSAGE lpNewMessage = nullptr;
			ULONG ulMessageStatus = NULL;

			auto hRes = WC_H(GetNextMessage(ulDir, &iNewItem, &ulMessageStatus, &lpNewMessage));
			if (lpNewMessage)
			{
				ULONG cValuesShow = 0;
				LPSPropValue lpspvaShow = nullptr;

				EC_H_GETPROPS_S(lpNewMessage->GetProps(
					LPSPropTagArray(&sptaShowForm), // property tag array
					fMapiUnicode, // flags
					&cValuesShow, // Count of values returned
					&lpspvaShow)); // Values returned
				if (lpspvaShow)
				{
					LPPERSISTMESSAGE lpNewPersistMessage = nullptr;

					if (m_lpMapiFormAdviseSink)
					{
						output::DebugPrint(
							output::DBGFormViewer,
							L"Calling OnActivateNext: szClass = %hs, ulStatus = 0x%X, ulFlags = 0x%X\n",
							lpspvaShow[ePR_MESSAGE_CLASS_A].Value.lpszA,
							ulMessageStatus,
							lpspvaShow[ePR_MESSAGE_FLAGS].Value.ul);

						hRes = WC_MAPI(m_lpMapiFormAdviseSink->OnActivateNext(
							lpspvaShow[ePR_MESSAGE_CLASS_A].Value.lpszA,
							ulMessageStatus, // message status
							lpspvaShow[ePR_MESSAGE_FLAGS].Value.ul, // message flags
							&lpNewPersistMessage));
					}
					else
					{
						hRes = S_FALSE;
					}

					if (hRes == S_OK) // we can handle the message ourselves
					{
						if (lpNewPersistMessage)
						{
							output::DebugPrintEx(
								output::DBGFormViewer,
								CLASS,
								L"ActivateNext",
								L"Got new persist from OnActivateNext\n");

							EC_H_S(OpenMessageNonModal(
								m_hwndParent,
								m_lpMDB,
								m_lpMAPISession,
								m_lpFolder,
								m_lpContentsTableListCtrl,
								iNewItem,
								lpNewMessage,
								EXCHIVERB_OPEN,
								lpRect));
						}
						else
						{
							error::ErrDialog(__FILE__, __LINE__, IDS_EDACTIVATENEXT);
						}
					}
					// We have to load the form from scratch
					else if (hRes == S_FALSE)
					{
						output::DebugPrintEx(
							output::DBGFormViewer,
							CLASS,
							L"ActivateNext",
							L"Didn't get new persist from OnActivateNext\n");
						// we're going to return S_FALSE, which will shut us down, so we can spin a whole new site
						// we don't need to clean up this site since the shutdown will do it for us
						// BTW - it might be more efficient to in-line this code and eliminate a GetProps call
						EC_H_S(OpenMessageNonModal(
							m_hwndParent,
							m_lpMDB,
							m_lpMAPISession,
							m_lpFolder,
							m_lpContentsTableListCtrl,
							iNewItem,
							lpNewMessage,
							EXCHIVERB_OPEN,
							lpRect));
					}
					else
						CHECKHRESMSG(hRes, IDS_ONACTIVENEXTFAILED);

					if (lpNewPersistMessage) lpNewPersistMessage->Release();
					MAPIFreeBuffer(lpspvaShow);
				}

				lpNewMessage->Release();
			}

			return S_FALSE;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetPrintSetup(ULONG ulFlags, LPFORMPRINTSETUP* /*lppFormPrintSetup*/)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetPrintSetup", L"ulFlags = 0x%08X\n", ulFlags);
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetSaveStream(ULONG* /*pulFlags*/, ULONG* /*pulFormat*/, LPSTREAM* /*ppstm*/)
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetSaveStream", L"\n");
			return MAPI_E_NO_SUPPORT;
		}

		STDMETHODIMP CMyMAPIFormViewer::GetViewStatus(LPULONG lpulStatus)
		{
			const auto hRes = S_OK;

			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetViewStatus", L"\n");
			*lpulStatus = NULL;
			*lpulStatus |= VCSTATUS_INTERACTIVE;

			if (m_lpContentsTableListCtrl)
			{
				if (-1 != m_lpContentsTableListCtrl->GetNextItem(m_iItem, LVNI_BELOW))
				{
					*lpulStatus |= VCSTATUS_NEXT;
				}

				if (-1 != m_lpContentsTableListCtrl->GetNextItem(m_iItem, LVNI_ABOVE))
				{
					*lpulStatus |= VCSTATUS_PREV;
				}
			}

			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetViewStatus", L"Returning 0x%X\n", *lpulStatus);
			return hRes;
		}

		_Check_return_ HRESULT CMyMAPIFormViewer::GetNextMessage(
			ULONG ulDir,
			_Out_ int* piNewItem,
			_Out_ ULONG* pulStatus,
			_Deref_out_opt_ LPMESSAGE* ppMessage) const
		{
			output::DebugPrintEx(output::DBGFormViewer, CLASS, L"GetNextMessage", L"ulDir = 0x%X\n", ulDir);
			auto hRes = S_OK;

			*piNewItem = NULL;
			*pulStatus = NULL;
			*ppMessage = nullptr;

			// Without a view list control, we can't do 'next'
			if (!m_lpContentsTableListCtrl || !m_lpMDB) return MAPI_E_INVALID_PARAMETER;
			if (m_iItem == -1) return S_FALSE;

			if (ulDir & VCDIR_NEXT)
			{
				*piNewItem = m_lpContentsTableListCtrl->GetNextItem(m_iItem, LVNI_BELOW);
			}
			else if (ulDir & VCDIR_PREV)
			{
				*piNewItem = m_lpContentsTableListCtrl->GetNextItem(m_iItem, LVNI_ABOVE);
			}

			if (*piNewItem == -1)
			{
				hRes = S_FALSE;
			}
			else
			{
				const auto lpData = m_lpContentsTableListCtrl->GetSortListData(*piNewItem);
				if (lpData && lpData->Contents())
				{
					const auto lpEID = lpData->Contents()->m_lpEntryID;
					if (lpEID)
					{
						*ppMessage = mapi::CallOpenEntry<LPMESSAGE>(
							m_lpMDB,
							nullptr,
							nullptr,
							nullptr,
							lpEID->cb,
							reinterpret_cast<LPENTRYID>(lpEID->lpb),
							nullptr,
							MAPI_BEST_ACCESS,
							nullptr);

						if (SUCCEEDED(hRes))
						{

							hRes = EC_MAPI(m_lpFolder->GetMessageStatus(
								lpEID->cb, reinterpret_cast<LPENTRYID>(lpEID->lpb), 0, pulStatus));
						}
					}
				}
			}

			if (FAILED(hRes))
			{
				*piNewItem = -1;
				*pulStatus = NULL;
				if (*ppMessage)
				{
					(*ppMessage)->Release();
					*ppMessage = nullptr;
				}
			}

			return hRes;
		}
	} // namespace mapiui
} // namespace mapi