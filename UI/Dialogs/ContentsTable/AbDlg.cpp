// Displays the contents of a single address book

#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/ABDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/columnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <core/mapi/mapiABFunctions.h>
#include <core/mapi/mapiProgress.h>
#include <core/mapi/cache/globalCache.h>
#include <UI/addinui.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CAbDlg";

	CAbDlg::CAbDlg(_In_ ui::CParentWnd* pParentWnd, _In_ cache::CMapiObjects* lpMapiObjects, _In_ LPMAPIPROP lpAbCont)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_AB,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  lpAbCont,
			  nullptr,
			  &columns::sptABCols.tags,
			  columns::ABColumns,
			  IDR_MENU_AB_VIEW_POPUP,
			  MENU_CONTEXT_AB_CONTENTS)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpAbCont = mapi::safe_cast<LPABCONT>(lpAbCont);

		m_bIsAB = true;

		CAbDlg::CreateDialogAndMenu(IDR_MENU_AB_VIEW);
	}

	CAbDlg::~CAbDlg()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpAbCont) m_lpAbCont->Release();
	}

	BEGIN_MESSAGE_MAP(CAbDlg, CContentsTableDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_DISPLAYDETAILS, OnDisplayDetails)
	ON_COMMAND(ID_OPENCONTACT, OnOpenContact)
	ON_COMMAND(ID_OPENMANAGER, OnOpenManager)
	ON_COMMAND(ID_OPENOWNER, OnOpenOwner)
	ON_COMMAND(ID_SEARCHFORNAME, OnCreatePropertyStringRestriction)
	END_MESSAGE_MAP()

	void CAbDlg::CreateDialogAndMenu(UINT nIDMenuResource)
	{
		output::DebugPrintEx(DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
		CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

		ui::UpdateMenuString(m_hWnd, ID_CREATEPROPERTYSTRINGRESTRICTION, IDS_ABRESMENU);
	}

	void CAbDlg::OnInitMenu(_In_ CMenu* pMenu)
	{
		if (pMenu && m_lpContentsTableListCtrl)
		{
			const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();
			pMenu->EnableMenuItem(ID_PASTE, DIM(ulStatus & BUFFER_ABENTRIES));

			const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_COPY, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_DISPLAYDETAILS, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_OPENCONTACT, DIMMSNOK(iNumSel));
			pMenu->EnableMenuItem(ID_OPENMANAGER, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_OPENOWNER, DIMMSOK(iNumSel));
		}

		CContentsTableDlg::OnInitMenu(pMenu);
	}

	void CAbDlg::OnDisplayDetails()
	{
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnDisplayDetails", L"displaying Address Book entry details\n");

		if (!m_lpMapiObjects) return;
		auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (lpAddrBook)
		{
			const auto lpEIDs = m_lpContentsTableListCtrl->GetSelectedItemEIDs();
			if (lpEIDs && lpEIDs->cValues && lpEIDs->lpbin)
			{
				for (ULONG i = 0; i < lpEIDs->cValues; i++)
				{
					auto ulUIParam = reinterpret_cast<ULONG_PTR>(static_cast<void*>(m_hWnd));

					// Have to pass DIALOG_MODAL according to
					// http://support.microsoft.com/kb/171637
					const auto hRes = EC_H_CANCEL(lpAddrBook->Details(
						&ulUIParam,
						NULL,
						NULL,
						lpEIDs->lpbin[i].cb,
						reinterpret_cast<LPENTRYID>(lpEIDs->lpbin[i].lpb),
						NULL,
						NULL,
						NULL,
						DIALOG_MODAL)); // API doesn't like Unicode
					if (lpEIDs->cValues > i + 1 && bShouldCancel(this, hRes)) break;
				}
			}

			MAPIFreeBuffer(lpEIDs);
		}
	}

	void CAbDlg::OnOpenContact()
	{
		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpPropDisplay) return;
		const auto lpSession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpSession) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto lpEntryList = m_lpContentsTableListCtrl->GetSelectedItemEIDs();
		if (lpEntryList && 1 == lpEntryList->cValues)
		{
			ULONG cb = 0;
			LPBYTE lpb = nullptr;
			if (mapi::UnwrapContactEntryID(lpEntryList->lpbin[0].cb, lpEntryList->lpbin[0].lpb, &cb, &lpb))
			{
				auto lpProp = mapi::CallOpenEntry<LPMAPIPROP>(
					nullptr, nullptr, nullptr, lpSession, cb, reinterpret_cast<LPENTRYID>(lpb), nullptr, NULL, nullptr);
				if (lpProp)
				{
					EC_H_S(DisplayObject(lpProp, NULL, otDefault, this));
					if (lpProp) lpProp->Release();
				}
			}
		}

		MAPIFreeBuffer(lpEntryList);
	}

	void CAbDlg::OnOpenManager()
	{
		auto iItem = -1;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		do
		{
			auto lpMailUser = m_lpContentsTableListCtrl->OpenNextSelectedItemProp(&iItem, mfcmapiREQUEST_MODIFY);

			if (lpMailUser)
			{
				EC_H_S(DisplayTable(
					lpMailUser,
					PR_EMS_AB_MANAGER_O,
					otDefault, // oType,
					this));

				lpMailUser->Release();
			}
		} while (iItem != -1);
	}

	void CAbDlg::OnOpenOwner()
	{
		auto iItem = -1;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		do
		{
			auto lpMailUser = m_lpContentsTableListCtrl->OpenNextSelectedItemProp(&iItem, mfcmapiREQUEST_MODIFY);

			if (lpMailUser)
			{
				EC_H_S(DisplayTable(
					lpMailUser,
					PR_EMS_AB_OWNER_O,
					otDefault, // oType,
					this));

				lpMailUser->Release();
			}
		} while (iItem != -1);
	}

	void CAbDlg::OnDeleteSelectedItem()
	{
		if (!m_lpAbCont) return;

		editor::CEditor Query(
			this, IDS_DELETEABENTRY, IDS_DELETEABENTRYPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		if (!Query.DisplayDialog()) return;

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnDeleteSelectedItem", L"deleting address Book entries\n");
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto lpEIDs = m_lpContentsTableListCtrl->GetSelectedItemEIDs();
		EC_MAPI_S(m_lpAbCont->DeleteEntries(lpEIDs, NULL));
		MAPIFreeBuffer(lpEIDs);
	}

	void CAbDlg::HandleCopy()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
		if (!m_lpContentsTableListCtrl) return;

		const auto lpEIDs = m_lpContentsTableListCtrl->GetSelectedItemEIDs();

		// CGlobalCache takes over ownership of lpEIDs - don't free now
		cache::CGlobalCache::getInstance().SetABEntriesToCopy(lpEIDs);
	}

	_Check_return_ bool CAbDlg::HandlePaste()
	{
		if (CBaseDialog::HandlePaste()) return true;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"pasting address Book entries\n");
		if (!m_lpAbCont) return false;

		const auto lpEIDs = cache::CGlobalCache::getInstance().GetABEntriesToCopy();

		if (lpEIDs)
		{
			editor::CEditor MyData(
				this, IDS_CALLCOPYENTRIES, IDS_CALLCOPYENTRIESPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLAGS, false));
			MyData.SetHex(0, CREATE_CHECK_DUP_STRICT);

			if (MyData.DisplayDialog())
			{
				auto lpProgress = mapi::mapiui::GetMAPIProgress(L"IABContainer::CopyEntries", m_hWnd); // STRING_OK

				EC_MAPI_S(m_lpAbCont->CopyEntries(
					lpEIDs, lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL, lpProgress, MyData.GetHex(0)));

				if (lpProgress) lpProgress->Release();
			}

			return true; // handled pasted
		}

		return false; // did not handle paste
	}

	void CAbDlg::OnCreatePropertyStringRestriction()
	{
		if (!m_lpContentsTableListCtrl) return;

		editor::CEditor MyData(
			this, IDS_SEARCHCRITERIA, IDS_ABSEARCHCRITERIAPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_NAME, false));

		if (!MyData.DisplayDialog()) return;

		// Allocate and create our SRestriction
		const auto lpRes = mapi::ab::CreateANRRestriction(PR_ANR_W, MyData.GetStringW(0), nullptr);
		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
	}

	void CAbDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER /*lpContainer*/)
	{
		if (lpParams)
		{
			lpParams->lpAbCont = m_lpAbCont;
			lpParams->lpMailUser = mapi::safe_cast<LPMAILUSER>(lpMAPIProp);
		}

		ui::addinui::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpMailUser)
		{
			lpParams->lpMailUser->Release();
			lpParams->lpMailUser = nullptr;
		}
	}
} // namespace dialog