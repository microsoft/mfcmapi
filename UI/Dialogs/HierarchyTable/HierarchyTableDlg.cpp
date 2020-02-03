#include <StdAfx.h>
#include <UI/Dialogs/HierarchyTable/HierarchyTableDlg.h>
#include <UI/Controls/StyleTree/HierarchyTableTreeCtrl.h>
#include <UI/FakeSplitter.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/RestrictEditor.h>
#include <UI/addinui.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CHierarchyTableDlg";

	CHierarchyTableDlg::CHierarchyTableDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		UINT uidTitle,
		_In_opt_ LPMAPIPROP lpRootContainer,
		ULONG nIDContextMenu,
		ULONG ulAddInContext)
		: CBaseDialog(pParentWnd, lpMapiObjects, ulAddInContext)
	{
		TRACE_CONSTRUCTOR(CLASS);
		if (NULL != uidTitle)
		{
			m_szTitle = strings::loadstring(uidTitle);
		}
		else
		{
			m_szTitle = strings::loadstring(IDS_TABLEASHIERARCHY);
		}

		m_nIDContextMenu = nIDContextMenu;

		m_ulDisplayFlags = dfNormal;
		m_lpContainer = mapi::safe_cast<LPMAPICONTAINER>(lpRootContainer);
	}

	CHierarchyTableDlg::~CHierarchyTableDlg()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpContainer) m_lpContainer->Release();
	}

	BEGIN_MESSAGE_MAP(CHierarchyTableDlg, CBaseDialog)
	ON_COMMAND(ID_DISPLAYSELECTEDITEM, OnDisplayItem)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	ON_COMMAND(ID_DISPLAYHIERARCHYTABLE, OnDisplayHierarchyTable)
	ON_COMMAND(ID_EDITSEARCHCRITERIA, OnEditSearchCriteria)
	END_MESSAGE_MAP()

	void CHierarchyTableDlg::OnInitMenu(_In_ CMenu* pMenu)
	{
		if (pMenu)
		{
			if (m_lpHierarchyTableTreeCtrl)
			{
				const auto bItemSelected = m_lpHierarchyTableTreeCtrl.IsItemSelected();
				pMenu->EnableMenuItem(ID_DISPLAYSELECTEDITEM, DIM(bItemSelected));
				pMenu->EnableMenuItem(ID_DISPLAYHIERARCHYTABLE, DIM(bItemSelected));
				pMenu->EnableMenuItem(ID_EDITSEARCHCRITERIA, DIM(bItemSelected));
				for (ULONG ulMenu = ID_ADDINMENU; ulMenu < ID_ADDINMENU + m_ulAddInMenuItems; ulMenu++)
				{
					const auto lpAddInMenu = ui::addinui::GetAddinMenuItem(m_hWnd, ulMenu);
					if (!lpAddInMenu) continue;

					const auto ulFlags = lpAddInMenu->ulFlags;
					UINT uiEnable = MF_ENABLED;

					if (ulFlags & MENU_FLAGS_MULTISELECT && !bItemSelected) uiEnable = MF_GRAYED;
					if (ulFlags & MENU_FLAGS_SINGLESELECT && !bItemSelected) uiEnable = MF_GRAYED;
					EnableAddInMenus(pMenu->m_hMenu, ulMenu, lpAddInMenu, uiEnable);
				}
			}
		}

		CBaseDialog::OnInitMenu(pMenu);
	}

	// Remove previous container and set new one. Will addref container.
	void CHierarchyTableDlg::SetRootContainer(LPUNKNOWN container)
	{
		if (m_lpContainer) m_lpContainer->Release();
		m_lpContainer = mapi::safe_cast<LPMAPICONTAINER>(container);
	}

	void CHierarchyTableDlg::OnCancel()
	{
		ShowWindow(SW_HIDE);
		CBaseDialog::OnCancel();
	}

	void CHierarchyTableDlg::OnDisplayItem()
	{
		auto lpMAPIContainer = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);
		if (!lpMAPIContainer)
		{
			WARNHRESMSG(MAPI_E_NOT_FOUND, IDS_NOITEMSELECTED);
			return;
		}

		EC_H_S(DisplayObject(lpMAPIContainer, NULL, otContents, this));

		lpMAPIContainer->Release();
	}

	void CHierarchyTableDlg::OnDisplayHierarchyTable()
	{
		LPMAPITABLE lpMAPITable = nullptr;

		if (!m_lpHierarchyTableTreeCtrl) return;

		auto lpContainer = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);

		if (lpContainer)
		{
			editor::CEditor MyData(
				this,
				IDS_DISPLAYHIEARCHYTABLE,
				IDS_DISPLAYHIEARCHYTABLEPROMPT,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.AddPane(viewpane::CheckPane::Create(0, IDS_CONVENIENTDEPTH, false, false));

			if (MyData.DisplayDialog())
			{
				EC_MAPI_S(lpContainer->GetHierarchyTable(
					MyData.GetCheck(0) ? CONVENIENT_DEPTH : 0 | fMapiUnicode, &lpMAPITable));

				if (lpMAPITable)
				{
					EC_H_S(DisplayTable(lpMAPITable, otHierarchy, this));
					lpMAPITable->Release();
				}
			}

			lpContainer->Release();
		}
	}

	void CHierarchyTableDlg::OnEditSearchCriteria()
	{
		if (!m_lpHierarchyTableTreeCtrl) return;

		// Find the highlighted item
		auto container = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);
		if (!container) return;
		auto lpMAPIFolder = mapi::safe_cast<LPMAPIFOLDER>(container);
		container->Release();

		if (lpMAPIFolder)
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic,
				CLASS,
				L"OnEditSearchCriteria",
				L"Calling GetSearchCriteria on %p.\n",
				lpMAPIFolder);

			LPSRestriction lpRes = nullptr;
			LPENTRYLIST lpEntryList = nullptr;
			ULONG ulSearchState = 0;

			const auto hRes =
				WC_MAPI(lpMAPIFolder->GetSearchCriteria(fMapiUnicode, &lpRes, &lpEntryList, &ulSearchState));
			if (hRes == MAPI_E_NOT_INITIALIZED)
			{
				output::DebugPrint(output::dbgLevel::Generic, L"No search criteria has been set on this folder.\n");
			}
			else
				CHECKHRESMSG(hRes, IDS_GETSEARCHCRITERIAFAILED);

			editor::CCriteriaEditor MyCriteria(this, lpRes, lpEntryList, ulSearchState);

			if (MyCriteria.DisplayDialog())
			{
				// make sure the user really wants to call SetSearchCriteria
				// hard to detect 'dirty' on this dialog so easier just to ask
				editor::CEditor MyYesNoDialog(
					this,
					IDS_CALLSETSEARCHCRITERIA,
					IDS_CALLSETSEARCHCRITERIAPROMPT,
					CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
				if (MyYesNoDialog.DisplayDialog())
				{
					// do the set search criteria
					const auto lpNewRes = MyCriteria.DetachModifiedSRestriction();
					const auto lpNewEntryList = MyCriteria.DetachModifiedEntryList();
					const auto ulSearchFlags = MyCriteria.GetSearchFlags();
					EC_MAPI_S(lpMAPIFolder->SetSearchCriteria(lpNewRes, lpNewEntryList, ulSearchFlags));
					MAPIFreeBuffer(lpNewRes);
					MAPIFreeBuffer(lpNewEntryList);
				}
			}

			lpMAPIFolder->Release();
		}
	}

	BOOL CHierarchyTableDlg::OnInitDialog()
	{
		const auto bRet = CBaseDialog::OnInitDialog();

		if (m_lpFakeSplitter)
		{
			m_lpHierarchyTableTreeCtrl.Create(
				&m_lpFakeSplitter, m_lpMapiObjects, this, m_ulDisplayFlags, m_nIDContextMenu);
			m_lpFakeSplitter.SetPaneOne(m_lpHierarchyTableTreeCtrl.GetSafeHwnd());

			m_lpFakeSplitter.SetPercent(0.25);
		}

		UpdateTitleBarText();

		return bRet;
	}

	void CHierarchyTableDlg::CreateDialogAndMenu(const UINT nIDMenuResource, const UINT uiClassMenuResource)
	{
		output::DebugPrintEx(
			output::dbgLevel::CreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
		CBaseDialog::CreateDialogAndMenu(nIDMenuResource, uiClassMenuResource, IDS_HIERARCHYTABLE);

		if (m_lpHierarchyTableTreeCtrl)
		{
			m_lpHierarchyTableTreeCtrl.LoadHierarchyTable(m_lpContainer);
		}
	}

	void CHierarchyTableDlg::OnRefreshView()
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnRefreshView", L"\n");
		if (m_lpHierarchyTableTreeCtrl) m_lpHierarchyTableTreeCtrl.Refresh();
	}

	_Check_return_ bool CHierarchyTableDlg::HandleAddInMenu(WORD wMenuSelect)
	{
		if (wMenuSelect < ID_ADDINMENU || ID_ADDINMENU + m_ulAddInMenuItems < wMenuSelect) return false;
		if (!m_lpHierarchyTableTreeCtrl) return false;

		LPMAPICONTAINER lpContainer = nullptr;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto lpAddInMenu = ui::addinui::GetAddinMenuItem(m_hWnd, wMenuSelect);
		if (!lpAddInMenu) return false;

		const auto ulFlags = lpAddInMenu->ulFlags;

		const auto fRequestModify =
			ulFlags & MENU_FLAGS_REQUESTMODIFY ? mfcmapiREQUEST_MODIFY : mfcmapiDO_NOT_REQUEST_MODIFY;

		// Get the stuff we need for any case
		_AddInMenuParams MyAddInMenuParams = {nullptr};
		MyAddInMenuParams.lpAddInMenu = lpAddInMenu;
		MyAddInMenuParams.ulAddInContext = m_ulAddInContext;
		MyAddInMenuParams.hWndParent = m_hWnd;
		if (m_lpMapiObjects)
		{
			MyAddInMenuParams.lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
			MyAddInMenuParams.lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			MyAddInMenuParams.lpAdrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
		}

		if (m_lpPropDisplay)
		{
			m_lpPropDisplay->GetSelectedPropTag(&MyAddInMenuParams.ulPropTag);
		}

		// MENU_FLAGS_SINGLESELECT and MENU_FLAGS_MULTISELECT can't both be set, so we can ignore this case
		if (!(ulFlags & (MENU_FLAGS_SINGLESELECT | MENU_FLAGS_MULTISELECT)))
		{
			HandleAddInMenuSingle(&MyAddInMenuParams, nullptr, nullptr);
		}
		else
		{
			SRow MyRow = {0};

			// If we have a row to give, give it - it's free
			const auto lpData = m_lpHierarchyTableTreeCtrl.GetSelectedItemData();
			if (lpData)
			{
				MyRow.cValues = lpData->cSourceProps;
				MyRow.lpProps = lpData->lpSourceProps;
				MyAddInMenuParams.lpRow = &MyRow;
				MyAddInMenuParams.ulCurrentFlags |= MENU_FLAGS_ROW;
			}

			if (!(ulFlags & MENU_FLAGS_ROW))
			{
				lpContainer = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(fRequestModify);
			}

			HandleAddInMenuSingle(&MyAddInMenuParams, nullptr, lpContainer);
			if (lpContainer) lpContainer->Release();
		}
		return true;
	}

	void CHierarchyTableDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_opt_ LPMAPIPROP /*lpMAPIProp*/,
		_In_opt_ LPMAPICONTAINER /*lpContainer*/)
	{
		ui::addinui::InvokeAddInMenu(lpParams);
	}
} // namespace dialog