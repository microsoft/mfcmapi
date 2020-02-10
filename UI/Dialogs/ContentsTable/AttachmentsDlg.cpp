// Displays the attachment table for a message

#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/AttachmentsDlg.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <UI/FileDialogEx.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/columnTags.h>
#include <core/mapi/mapiProgress.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <core/sortlistdata/contentsData.h>
#include <core/mapi/cache/globalCache.h>
#include <UI/mapiui.h>
#include <UI/addinui.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFile.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CAttachmentsDlg";

	CAttachmentsDlg::CAttachmentsDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ LPMAPITABLE lpMAPITable,
		_In_ LPMAPIPROP lpMessage)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_ATTACHMENTS,
			  createDialogType::DO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  lpMAPITable,
			  &columns::sptATTACHCols.tags,
			  columns::ATTACHColumns,
			  IDR_MENU_ATTACHMENTS_POPUP,
			  MENU_CONTEXT_ATTACHMENT_TABLE)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpMessage = mapi::safe_cast<LPMESSAGE>(lpMessage);

		m_bDisplayAttachAsEmbeddedMessage = false;
		m_lpAttach = nullptr;
		m_ulAttachNum = static_cast<ULONG>(-1);

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_ATTACHMENTS);
	}

	CAttachmentsDlg::~CAttachmentsDlg()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpAttach) m_lpAttach->Release();
		if (m_lpMessage)
		{
			WC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			m_lpMessage->Release();
		}
	}

	BEGIN_MESSAGE_MAP(CAttachmentsDlg, CContentsTableDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYSELECTEDITEM, OnModifySelectedItem)
	ON_COMMAND(ID_SAVECHANGES, OnSaveChanges)
	ON_COMMAND(ID_SAVETOFILE, OnSaveToFile)
	ON_COMMAND(ID_VIEWEMBEDDEDMESSAGEPROPERTIES, OnViewEmbeddedMessageProps)
	ON_COMMAND(ID_ADDATTACHMENT, OnAddAttachment)
	END_MESSAGE_MAP()

	void CAttachmentsDlg::OnInitMenu(_In_ CMenu* pMenu)
	{
		if (pMenu)
		{
			if (m_lpContentsTableListCtrl)
			{
				const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
				const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();
				pMenu->EnableMenuItem(ID_PASTE, DIM(ulStatus & BUFFER_ATTACHMENTS));

				pMenu->EnableMenuItem(ID_COPY, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_MODIFYSELECTEDITEM, DIMMSOK(1 == iNumSel));
				pMenu->EnableMenuItem(ID_SAVETOFILE, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_DISPLAYSELECTEDITEM, DIM(1 == iNumSel));
			}

			pMenu->CheckMenuItem(ID_VIEWEMBEDDEDMESSAGEPROPERTIES, CHECK(m_bDisplayAttachAsEmbeddedMessage));
		}

		CContentsTableDlg::OnInitMenu(pMenu);
	}

	void CAttachmentsDlg::OnDisplayItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
		if (!m_lpAttach) return;

		const auto lpListData = m_lpContentsTableListCtrl->GetFirstSelectedItemData();
		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents && ATTACH_EMBEDDED_MSG == contents->m_ulAttachMethod)
			{
				auto lpMessage = OpenEmbeddedMessage();
				if (lpMessage)
				{
					WC_H_S(DisplayObject(lpMessage, MAPI_MESSAGE, objectType::otDefault, this));
					lpMessage->Release();
				}
			}
		}
	}

	_Check_return_ LPATTACH CAttachmentsDlg::OpenAttach(ULONG ulAttachNum) const
	{
		LPATTACH lpAttach = nullptr;

		const auto hRes = WC_MAPI(m_lpMessage->OpenAttach(ulAttachNum, NULL, MAPI_MODIFY, &lpAttach));
		if (hRes == MAPI_E_NO_ACCESS)
		{
			WC_MAPI_S(m_lpMessage->OpenAttach(ulAttachNum, NULL, MAPI_BEST_ACCESS, &lpAttach));
		}

		return lpAttach;
	}

	_Check_return_ LPMESSAGE CAttachmentsDlg::OpenEmbeddedMessage() const
	{
		if (!m_lpAttach) return nullptr;

		LPMESSAGE lpMessage = nullptr;
		auto hRes = WC_MAPI(m_lpAttach->OpenProperty(
			PR_ATTACH_DATA_OBJ,
			const_cast<LPIID>(&IID_IMessage),
			0,
			MAPI_MODIFY,
			reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
		if (hRes == MAPI_E_NO_ACCESS)
		{
			hRes = WC_MAPI(m_lpAttach->OpenProperty(
				PR_ATTACH_DATA_OBJ,
				const_cast<LPIID>(&IID_IMessage),
				0,
				MAPI_BEST_ACCESS,
				reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
		}

		if (hRes == MAPI_E_INTERFACE_NOT_SUPPORTED || hRes == MAPI_E_NOT_FOUND)
		{
			WARNHRESMSG(hRes, IDS_ATTNOTEMBEDDEDMSG);
		}

		return lpMessage;
	}

	_Check_return_ LPMAPIPROP CAttachmentsDlg::OpenItemProp(int iSelectedItem, modifyType /*bModify*/)
	{
		if (!m_lpContentsTableListCtrl) return nullptr;
		output::DebugPrintEx(
			output::dbgLevel::OpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		// Find the highlighted item AttachNum
		const auto lpListData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);

		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents)
			{
				const auto ulAttachNum = contents->m_ulAttachNum;
				const auto ulAttachMethod = contents->m_ulAttachMethod;

				// Check for matching cached attachment to avoid reopen
				if (ulAttachNum != m_ulAttachNum || !m_lpAttach)
				{
					if (m_lpAttach) m_lpAttach->Release();
					m_lpAttach = OpenAttach(ulAttachNum);
					m_ulAttachNum = static_cast<ULONG>(-1);
					if (m_lpAttach)
					{
						m_ulAttachNum = ulAttachNum;
					}
				}

				if (m_lpAttach && m_bDisplayAttachAsEmbeddedMessage && ATTACH_EMBEDDED_MSG == ulAttachMethod)
				{
					// Reopening an embedded message can fail
					// The view might be holding the embedded message we're trying to open, so we clear
					// it from the view to allow us to reopen it.
					// TODO: Consider caching our embedded message so this isn't necessary
					OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);
					return OpenEmbeddedMessage();
				}
				else
				{
					if (m_lpAttach) m_lpAttach->AddRef();
					return m_lpAttach;
				}
			}
		}

		return nullptr;
	}

	void CAttachmentsDlg::HandleCopy()
	{
		if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"HandleCopy", L"\n");
		if (!m_lpContentsTableListCtrl) return;

		const ULONG ulNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

		if (ulNumSelected && ulNumSelected < ULONG_MAX / sizeof(ULONG))
		{
			std::vector<ULONG> lpAttNumList;
			auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
			for (const auto& lpListData : items)
			{
				if (lpListData)
				{
					const auto contents = lpListData->cast<sortlistdata::contentsData>();
					if (contents)
					{
						lpAttNumList.push_back(contents->m_ulAttachNum);
					}
				}
			}

			cache::CGlobalCache::getInstance().SetAttachmentsToCopy(m_lpMessage, lpAttNumList);
		}
	}

	_Check_return_ bool CAttachmentsDlg::HandlePaste()
	{
		if (CBaseDialog::HandlePaste()) return true;

		if (!m_lpContentsTableListCtrl || !m_lpMessage) return false;
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"HandlePaste", L"\n");

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();
		if (!(ulStatus & BUFFER_ATTACHMENTS) || !(ulStatus & BUFFER_SOURCEPROPOBJ)) return false;

		auto lpAttNumList = cache::CGlobalCache::getInstance().GetAttachmentsToCopy();
		auto lpSourceMessage = mapi::safe_cast<LPMESSAGE>(cache::CGlobalCache::getInstance().GetSourcePropObject());

		if (!lpAttNumList.empty() && lpSourceMessage)
		{
			// Go through each attachment and copy it
			for (const auto& ulAtt : lpAttNumList)
			{
				LPATTACH lpAttSrc = nullptr;
				LPATTACH lpAttDst = nullptr;
				LPSPropProblemArray lpProblems = nullptr;

				// Open the attachment source
				EC_MAPI_S(lpSourceMessage->OpenAttach(ulAtt, NULL, MAPI_DEFERRED_ERRORS, &lpAttSrc));
				if (lpAttSrc)
				{
					ULONG ulAttNum = NULL;
					// Create the attachment destination
					EC_MAPI_S(m_lpMessage->CreateAttach(NULL, MAPI_DEFERRED_ERRORS, &ulAttNum, &lpAttDst));
					if (lpAttDst)
					{
						auto lpProgress = mapi::mapiui::GetMAPIProgress(L"IAttach::CopyTo", m_hWnd); // STRING_OK

						// Copy from source to destination
						EC_MAPI_S(lpAttSrc->CopyTo(
							0,
							NULL,
							nullptr,
							lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
							lpProgress,
							const_cast<LPIID>(&IID_IAttachment),
							lpAttDst,
							lpProgress ? MAPI_DIALOG : 0,
							&lpProblems));

						if (lpProgress) lpProgress->Release();

						EC_PROBLEMARRAY(lpProblems);
						MAPIFreeBuffer(lpProblems);
					}
				}

				if (lpAttSrc) lpAttSrc->Release();
				lpAttSrc = nullptr;

				if (lpAttDst)
				{
					EC_MAPI_S(lpAttDst->SaveChanges(KEEP_OPEN_READWRITE));
					lpAttDst->Release();
					lpAttDst = nullptr;
				}

				// If we failed on one pass, try the rest
			}

			EC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			OnRefreshView(); // Update the view since we don't have notifications here.
		}

		if (lpSourceMessage) lpSourceMessage->Release();
		return true;
	}

	void CAttachmentsDlg::OnDeleteSelectedItem()
	{
		if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		std::vector<int> attachnums;
		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (lpListData)
			{
				const auto contents = lpListData->cast<sortlistdata::contentsData>();
				if (contents)
				{
					attachnums.push_back(contents->m_ulAttachNum);
				}
			}
		}

		for (const auto& attachnum : attachnums)
		{
			output::DebugPrintEx(
				output::dbgLevel::DeleteSelectedItem,
				CLASS,
				L"OnDeleteSelectedItem",
				L"Deleting attachment 0x%08X\n",
				attachnum);
			auto lpProgress = mapi::mapiui::GetMAPIProgress(L"IMessage::DeleteAttach", m_hWnd); // STRING_OK
			EC_MAPI_S(m_lpMessage->DeleteAttach(
				attachnum,
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
				lpProgress,
				lpProgress ? ATTACH_DIALOG : 0));

			if (lpProgress) lpProgress->Release();
		}

		EC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		OnRefreshView(); // Update the view since we don't have notifications here.
	}

	void CAttachmentsDlg::OnModifySelectedItem()
	{
		if (m_lpAttach)
		{
			EC_MAPI_S(m_lpAttach->SaveChanges(KEEP_OPEN_READWRITE));
		}
	}

	void CAttachmentsDlg::OnSaveChanges()
	{
		if (m_lpMessage)
		{
			EC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		}
	}

	void CAttachmentsDlg::OnSaveToFile()
	{
		auto hRes = S_OK;
		LPATTACH lpAttach = nullptr;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl || !m_lpMessage) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			// Find the highlighted item AttachNum
			if (S_OK != hRes)
			{
				if (bShouldCancel(this, hRes)) break;
				hRes = S_OK;
			}

			if (lpListData)
			{
				const auto contents = lpListData->cast<sortlistdata::contentsData>();
				if (contents)
				{
					const auto ulAttachNum = contents->m_ulAttachNum;

					EC_MAPI_S(m_lpMessage->OpenAttach(
						ulAttachNum, NULL, MAPI_BEST_ACCESS, static_cast<LPATTACH*>(&lpAttach)));

					if (lpAttach)
					{
						hRes = WC_H(ui::mapiui::WriteAttachmentToFile(lpAttach, m_hWnd));

						lpAttach->Release();
						lpAttach = nullptr;
					}
				}
			}
		}
	}

	void CAttachmentsDlg::OnViewEmbeddedMessageProps()
	{
		m_bDisplayAttachAsEmbeddedMessage = !m_bDisplayAttachAsEmbeddedMessage;
		OnRefreshView();
	}

	void CAttachmentsDlg::OnAddAttachment()
	{
		auto szAttachName = file::CFileDialogExW::OpenFile(
			strings::emptystring, strings::emptystring, NULL, strings::loadstring(IDS_ALLFILES));
		if (!szAttachName.empty())
		{
			LPATTACH lpAttachment = nullptr;
			ULONG ulAttachNum = 0;

			auto hRes = EC_MAPI(m_lpMessage->CreateAttach(NULL, NULL, &ulAttachNum, &lpAttachment));

			if (SUCCEEDED(hRes) && lpAttachment)
			{
				SPropValue spvAttach[4];
				spvAttach[0].ulPropTag = PR_ATTACH_METHOD;
				spvAttach[0].Value.l = ATTACH_BY_VALUE;
				spvAttach[1].ulPropTag = PR_RENDERING_POSITION;
				spvAttach[1].Value.l = -1;
				spvAttach[2].ulPropTag = PR_ATTACH_FILENAME_W;
				spvAttach[2].Value.lpszW = LPWSTR(szAttachName.c_str());
				spvAttach[3].ulPropTag = PR_DISPLAY_NAME_W;
				spvAttach[3].Value.lpszW = LPWSTR(szAttachName.c_str());

				hRes = EC_MAPI(lpAttachment->SetProps(_countof(spvAttach), spvAttach, NULL));
				if (SUCCEEDED(hRes))
				{
					LPSTREAM pStreamFile = nullptr;

					hRes = EC_MAPI(file::MyOpenStreamOnFile(
						MAPIAllocateBuffer, MAPIFreeBuffer, STGM_READ, szAttachName, &pStreamFile));
					if (SUCCEEDED(hRes) && pStreamFile)
					{
						LPSTREAM pStreamAtt = nullptr;
						STATSTG StatInfo = {nullptr};

						hRes = EC_MAPI(lpAttachment->OpenProperty(
							PR_ATTACH_DATA_BIN,
							&IID_IStream,
							0,
							MAPI_MODIFY | MAPI_CREATE,
							reinterpret_cast<LPUNKNOWN*>(&pStreamAtt)));
						if (SUCCEEDED(hRes) && pStreamAtt)
						{
							hRes = EC_MAPI(pStreamFile->Stat(&StatInfo, STATFLAG_NONAME));

							if (SUCCEEDED(hRes))
							{
								hRes = EC_MAPI(pStreamFile->CopyTo(pStreamAtt, StatInfo.cbSize, NULL, NULL));
							}

							if (SUCCEEDED(hRes))
							{
								hRes = EC_MAPI(pStreamAtt->Commit(STGC_DEFAULT));
							}

							if (SUCCEEDED(hRes))
							{
								hRes = EC_MAPI(lpAttachment->SaveChanges(KEEP_OPEN_READWRITE));
							}

							if (SUCCEEDED(hRes))
							{
								EC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
							}
						}

						if (pStreamAtt) pStreamAtt->Release();
					}

					if (pStreamFile) pStreamFile->Release();
				}
			}

			if (lpAttachment) lpAttachment->Release();

			OnRefreshView();
		}
	}

	void CAttachmentsDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER /*lpContainer*/)
	{
		if (lpParams)
		{
			lpParams->lpMessage = m_lpMessage;
			lpParams->lpAttach = mapi::safe_cast<LPATTACH>(lpMAPIProp);
		}

		ui::addinui::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpAttach)
		{
			lpParams->lpAttach->Release();
			lpParams->lpAttach = nullptr;
		}
	}
} // namespace dialog
