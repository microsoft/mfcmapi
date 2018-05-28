// Displays the attachment table for a message

#include "StdAfx.h"
#include "AttachmentsDlg.h"
#include <UI/Controls/ContentsTableListCtrl.h>
#include <IO/File.h>
#include <UI/FileDialogEx.h>
#include <MAPI/MapiObjects.h>
#include <MAPI/ColumnTags.h>
#include <MAPI/MAPIProgress.h>
#include <UI/MFCUtilityFunctions.h>
#include "ImportProcs.h"
#include <UI/Controls/SortList/ContentsData.h>
#include <MAPI/GlobalCache.h>
#include <Interpret/InterpretProp.h>

namespace dialog
{
	static std::wstring CLASS = L"CAttachmentsDlg";

	CAttachmentsDlg::CAttachmentsDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPMAPITABLE lpMAPITable,
		_In_ LPMESSAGE lpMessage
	) :
		CContentsTableDlg(
			pParentWnd,
			lpMapiObjects,
			IDS_ATTACHMENTS,
			mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			lpMAPITable,
			LPSPropTagArray(&sptATTACHCols),
			ATTACHColumns,
			IDR_MENU_ATTACHMENTS_POPUP,
			MENU_CONTEXT_ATTACHMENT_TABLE)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpMessage = lpMessage;
		if (m_lpMessage) m_lpMessage->AddRef();

		m_bDisplayAttachAsEmbeddedMessage = false;
		m_lpAttach = nullptr;
		m_ulAttachNum = static_cast<ULONG>(-1);

		CreateDialogAndMenu(IDR_MENU_ATTACHMENTS);
	}

	CAttachmentsDlg::~CAttachmentsDlg()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpAttach) m_lpAttach->Release();
		if (m_lpMessage)
		{
			auto hRes = S_OK;

			WC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
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
				const auto ulStatus = CGlobalCache::getInstance().GetBufferStatus();
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
		if (lpListData && lpListData->Contents() && ATTACH_EMBEDDED_MSG == lpListData->Contents()->m_ulAttachMethod)
		{
			auto hRes = S_OK;
			auto lpMessage = OpenEmbeddedMessage();
			if (lpMessage)
			{
				WC_H(DisplayObject(lpMessage, MAPI_MESSAGE, otDefault, this));
				lpMessage->Release();
			}
		}
	}

	_Check_return_ LPATTACH CAttachmentsDlg::OpenAttach(ULONG ulAttachNum) const
	{
		auto hRes = S_OK;
		LPATTACH lpAttach = nullptr;

		WC_MAPI(m_lpMessage->OpenAttach(
			ulAttachNum,
			NULL,
			MAPI_MODIFY,
			&lpAttach));
		if (MAPI_E_NO_ACCESS == hRes)
		{
			hRes = S_OK;
			WC_MAPI(m_lpMessage->OpenAttach(
				ulAttachNum,
				NULL,
				MAPI_BEST_ACCESS,
				&lpAttach));
		}

		return lpAttach;
	}

	_Check_return_ LPMESSAGE CAttachmentsDlg::OpenEmbeddedMessage() const
	{
		if (!m_lpAttach) return nullptr;
		auto hRes = S_OK;

		LPMESSAGE lpMessage = nullptr;
		WC_MAPI(m_lpAttach->OpenProperty(
			PR_ATTACH_DATA_OBJ,
			const_cast<LPIID>(&IID_IMessage),
			0,
			MAPI_MODIFY,
			reinterpret_cast<LPUNKNOWN *>(&lpMessage)));
		if (hRes == MAPI_E_NO_ACCESS)
		{
			hRes = S_OK;
			WC_MAPI(m_lpAttach->OpenProperty(
				PR_ATTACH_DATA_OBJ,
				const_cast<LPIID>(&IID_IMessage),
				0,
				MAPI_BEST_ACCESS,
				reinterpret_cast<LPUNKNOWN *>(&lpMessage)));
		}

		if (hRes == MAPI_E_INTERFACE_NOT_SUPPORTED ||
			hRes == MAPI_E_NOT_FOUND)
		{
			WARNHRESMSG(hRes, IDS_ATTNOTEMBEDDEDMSG);
		}

		return lpMessage;
	}

	_Check_return_ HRESULT CAttachmentsDlg::OpenItemProp(
		int iSelectedItem,
		__mfcmapiModifyEnum /*bModify*/,
		_Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
	{
		const auto hRes = S_OK;

		DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		if (!m_lpContentsTableListCtrl || !lppMAPIProp) return MAPI_E_INVALID_PARAMETER;

		*lppMAPIProp = nullptr;

		// Find the highlighted item AttachNum
		const auto lpListData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);

		if (lpListData && lpListData->Contents())
		{
			const auto ulAttachNum = lpListData->Contents()->m_ulAttachNum;
			const auto ulAttachMethod = lpListData->Contents()->m_ulAttachMethod;

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
				*lppMAPIProp = OpenEmbeddedMessage();
			}
			else
			{
				*lppMAPIProp = m_lpAttach;
				if (*lppMAPIProp) (*lppMAPIProp)->AddRef();
			}
		}

		return hRes;
	}

	void CAttachmentsDlg::HandleCopy()
	{
		if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
		if (!m_lpContentsTableListCtrl) return;

		const ULONG ulNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

		if (ulNumSelected && ulNumSelected < ULONG_MAX / sizeof(ULONG))
		{
			std::vector<ULONG> lpAttNumList;
			auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
			for (const auto& lpListData : items)
			{
				if (lpListData && lpListData->Contents())
				{
					lpAttNumList.push_back(lpListData->Contents()->m_ulAttachNum);
				}
			}

			CGlobalCache::getInstance().SetAttachmentsToCopy(m_lpMessage, lpAttNumList);
		}
	}

	_Check_return_ bool CAttachmentsDlg::HandlePaste()
	{
		if (CBaseDialog::HandlePaste()) return true;

		if (!m_lpContentsTableListCtrl || !m_lpMessage) return false;
		DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");

		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto ulStatus = CGlobalCache::getInstance().GetBufferStatus();
		if (!(ulStatus & BUFFER_ATTACHMENTS) || !(ulStatus & BUFFER_SOURCEPROPOBJ)) return false;

		auto lpAttNumList = CGlobalCache::getInstance().GetAttachmentsToCopy();
		auto lpSourceMessage = dynamic_cast<LPMESSAGE>(CGlobalCache::getInstance().GetSourcePropObject());

		if (!lpAttNumList.empty() && lpSourceMessage)
		{
			// Go through each attachment and copy it
			for (const auto& ulAtt : lpAttNumList)
			{
				LPATTACH lpAttSrc = nullptr;
				LPATTACH lpAttDst = nullptr;
				LPSPropProblemArray lpProblems = nullptr;

				// Open the attachment source
				EC_MAPI(lpSourceMessage->OpenAttach(
					ulAtt,
					NULL,
					MAPI_DEFERRED_ERRORS,
					&lpAttSrc));

				if (lpAttSrc)
				{
					ULONG ulAttNum = NULL;
					// Create the attachment destination
					EC_MAPI(m_lpMessage->CreateAttach(NULL, MAPI_DEFERRED_ERRORS, &ulAttNum, &lpAttDst));
					if (lpAttDst)
					{
						LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IAttach::CopyTo", m_hWnd); // STRING_OK

						// Copy from source to destination
						EC_MAPI(lpAttSrc->CopyTo(
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
					EC_MAPI(lpAttDst->SaveChanges(KEEP_OPEN_READWRITE));
					lpAttDst->Release();
					lpAttDst = nullptr;
				}

				// If we failed on one pass, try the rest
				hRes = S_OK;
			}

			EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			OnRefreshView(); // Update the view since we don't have notifications here.
		}

		return true;
	}

	void CAttachmentsDlg::OnDeleteSelectedItem()
	{
		if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		std::vector<int> attachnums;
		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (lpListData && lpListData->Contents())
			{
				attachnums.push_back(lpListData->Contents()->m_ulAttachNum);
			}
		}

		for (const auto& attachnum : attachnums)
		{
			DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L"Deleting attachment 0x%08X\n", attachnum);
			LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMessage::DeleteAttach", m_hWnd); // STRING_OK
			EC_MAPI(m_lpMessage->DeleteAttach(
				attachnum,
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
				lpProgress,
				lpProgress ? ATTACH_DIALOG : 0));

			if (lpProgress)
				lpProgress->Release();
		}

		EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		OnRefreshView(); // Update the view since we don't have notifications here.
	}

	void CAttachmentsDlg::OnModifySelectedItem()
	{
		if (m_lpAttach)
		{
			auto hRes = S_OK;

			EC_MAPI(m_lpAttach->SaveChanges(KEEP_OPEN_READWRITE));
		}
	}

	void CAttachmentsDlg::OnSaveChanges()
	{
		if (m_lpMessage)
		{
			auto hRes = S_OK;

			EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
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

			if (lpListData && lpListData->Contents())
			{
				const auto ulAttachNum = lpListData->Contents()->m_ulAttachNum;

				EC_MAPI(m_lpMessage->OpenAttach(
					ulAttachNum,
					NULL,
					MAPI_BEST_ACCESS,
					static_cast<LPATTACH*>(&lpAttach)));

				if (lpAttach)
				{
					WC_H(file::WriteAttachmentToFile(lpAttach, m_hWnd));

					lpAttach->Release();
					lpAttach = nullptr;
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
		HRESULT hRes = 0;
		auto szAttachName = file::CFileDialogExW::OpenFile(
			strings::emptystring,
			strings::emptystring,
			NULL,
			strings::loadstring(IDS_ALLFILES));
		if (!szAttachName.empty())
		{
			LPATTACH lpAttachment = nullptr;
			ULONG ulAttachNum = 0;

			EC_MAPI(m_lpMessage->CreateAttach(NULL, NULL, &ulAttachNum, &lpAttachment));

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

				EC_MAPI(lpAttachment->SetProps(_countof(spvAttach), spvAttach, NULL));
				if (SUCCEEDED(hRes))
				{
					LPSTREAM pStreamFile = nullptr;

					EC_MAPI(MyOpenStreamOnFile(
						MAPIAllocateBuffer,
						MAPIFreeBuffer,
						STGM_READ,
						szAttachName,
						&pStreamFile));
					if (SUCCEEDED(hRes) && pStreamFile)
					{
						LPSTREAM pStreamAtt = nullptr;
						STATSTG StatInfo = { nullptr };

						EC_MAPI(lpAttachment->OpenProperty(
							PR_ATTACH_DATA_BIN,
							&IID_IStream,
							0,
							MAPI_MODIFY | MAPI_CREATE,
							reinterpret_cast<LPUNKNOWN *>(&pStreamAtt)));
						if (SUCCEEDED(hRes) && pStreamAtt)
						{
							EC_MAPI(pStreamFile->Stat(&StatInfo, STATFLAG_NONAME));
							EC_MAPI(pStreamFile->CopyTo(pStreamAtt, StatInfo.cbSize, NULL, NULL));
							EC_MAPI(pStreamAtt->Commit(STGC_DEFAULT));
							EC_MAPI(lpAttachment->SaveChanges(KEEP_OPEN_READWRITE));
							EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
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
			lpParams->lpAttach = dynamic_cast<LPATTACH>(lpMAPIProp); // OpenItemProp returns LPATTACH
		}

		InvokeAddInMenu(lpParams);
	}
}