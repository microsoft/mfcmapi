// Displays the ACL table for an item

#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/AclDlg.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/columnTags.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <core/interpret/flags.h>
#include <core/mapi/mapiMemory.h>
#include <UI/addinui.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CAclDlg";

#define ACL_INCLUDE_ID 0x00000001
#define ACL_INCLUDE_OTHER 0x00000002

	CAclDlg::CAclDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ LPEXCHANGEMODIFYTABLE lpExchTbl,
		bool fFreeBusyVisible)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  fFreeBusyVisible ? IDS_ACLFBTABLE : IDS_ACLTABLE,
			  createDialogType::DO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  nullptr,
			  &columns::sptACLCols.tags,
			  columns::ACLColumns,
			  IDR_MENU_ACL_POPUP,
			  MENU_CONTEXT_ACL_TABLE)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpExchTbl = lpExchTbl;

		if (m_lpExchTbl) m_lpExchTbl->AddRef();

		if (fFreeBusyVisible)
			m_ulTableFlags = ACLTABLE_FREEBUSY;
		else
			m_ulTableFlags = 0;

		m_bIsAB = false;

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_ACL);

		CAclDlg::OnRefreshView();
	}

	CAclDlg::~CAclDlg()
	{
		TRACE_DESTRUCTOR(CLASS);

		if (m_lpExchTbl) m_lpExchTbl->Release();
	}

	BEGIN_MESSAGE_MAP(CAclDlg, CContentsTableDlg)
	ON_COMMAND(ID_ADDITEM, OnAddItem)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYSELECTEDITEM, OnModifySelectedItem)
	END_MESSAGE_MAP()

	_Check_return_ LPMAPIPROP CAclDlg::OpenItemProp(int /*iSelectedItem*/, modifyType /*bModify*/)
	{
		// Don't do anything because we don't want to override the properties that we have
		return nullptr;
	}

	void CAclDlg::OnInitMenu(_In_ CMenu* pMenu)
	{
		if (pMenu)
		{
			if (m_lpContentsTableListCtrl)
			{
				const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
				pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_MODIFYSELECTEDITEM, DIMMSOK(iNumSel));
			}
		}

		CContentsTableDlg::OnInitMenu(pMenu);
	}

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	void CAclDlg::OnRefreshView()
	{
		if (!m_lpExchTbl || !m_lpContentsTableListCtrl) return;

		if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnRefreshView", L"\n");

		if (m_lpExchTbl)
		{
			LPMAPITABLE lpMAPITable = nullptr;
			// Open a MAPI table on the Exchange table property. This table can be
			// read to determine what the Exchange table looks like.
			EC_MAPI_S(m_lpExchTbl->GetTable(m_ulTableFlags, &lpMAPITable));

			if (lpMAPITable)
			{
				m_lpContentsTableListCtrl->SetContentsTable(lpMAPITable, tableDisplayFlags::dfNormal, NULL);

				lpMAPITable->Release();
			}
		}
	}

	void CAclDlg::OnAddItem()
	{
		editor::CEditor MyData(this, IDS_ACLADDITEM, IDS_ACLADDITEMPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(flags::AllFlagsToString(PROP_ID(PR_MEMBER_RIGHTS), true));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_USEREID, false));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_MASKINHEX, false));
		MyData.SetHex(1, 0);

		if (!MyData.DisplayDialog())
		{
			output::DebugPrint(output::dbgLevel::Generic, L"OnAddItem cancelled.\n");
			return;
		}

		auto lpNewItem = mapi::allocate<LPROWLIST>(CbNewROWLIST(1));
		if (lpNewItem)
		{
			lpNewItem->cEntries = 1;
			lpNewItem->aEntries[0].ulRowFlags = ROW_ADD;
			lpNewItem->aEntries[0].cValues = 2;
			lpNewItem->aEntries[0].rgPropVals = nullptr;

			lpNewItem->aEntries[0].rgPropVals = mapi::allocate<LPSPropValue>(2 * sizeof(SPropValue), lpNewItem);
			if (lpNewItem->aEntries[0].rgPropVals)
			{
				const auto bin = MyData.GetBinary(0, false);

				lpNewItem->aEntries[0].rgPropVals[0].ulPropTag = PR_MEMBER_ENTRYID;
				lpNewItem->aEntries[0].rgPropVals[0].Value.bin =
					SBinary{static_cast<ULONG>(bin.size()), const_cast<BYTE*>(bin.data())};
				lpNewItem->aEntries[0].rgPropVals[1].ulPropTag = PR_MEMBER_RIGHTS;
				lpNewItem->aEntries[0].rgPropVals[1].Value.ul = MyData.GetHex(1);

				const auto hRes = EC_MAPI(m_lpExchTbl->ModifyTable(m_ulTableFlags, lpNewItem));
				MAPIFreeBuffer(lpNewItem);
				if (hRes == S_OK) OnRefreshView();
			}
		}
	}

	void CAclDlg::OnDeleteSelectedItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		LPROWLIST lpSelectedItems = nullptr;

		auto hRes = EC_H(GetSelectedItems(ACL_INCLUDE_ID, ROW_REMOVE, &lpSelectedItems));

		if (lpSelectedItems)
		{
			hRes = EC_MAPI(m_lpExchTbl->ModifyTable(m_ulTableFlags, lpSelectedItems));
			MAPIFreeBuffer(lpSelectedItems);
			if (hRes == S_OK) OnRefreshView();
		}
	}

	void CAclDlg::OnModifySelectedItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		LPROWLIST lpSelectedItems = nullptr;

		EC_H_S(GetSelectedItems(ACL_INCLUDE_ID | ACL_INCLUDE_OTHER, ROW_MODIFY, &lpSelectedItems));

		if (lpSelectedItems)
		{
			const auto hRes = EC_MAPI(m_lpExchTbl->ModifyTable(m_ulTableFlags, lpSelectedItems));
			MAPIFreeBuffer(lpSelectedItems);
			if (hRes == S_OK) OnRefreshView();
		}
	}

	_Check_return_ HRESULT CAclDlg::GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, _In_ LPROWLIST* lppRowList) const
	{
		if (!lppRowList || !m_lpContentsTableListCtrl) return MAPI_E_INVALID_PARAMETER;

		*lppRowList = nullptr;
		const int iNumItems = m_lpContentsTableListCtrl->GetSelectedCount();

		if (!iNumItems) return S_OK;
		if (iNumItems > MAXNewROWLIST) return MAPI_E_INVALID_PARAMETER;

		auto lpTempList = mapi::allocate<LPROWLIST>(CbNewROWLIST(iNumItems));
		if (lpTempList)
		{
			lpTempList->cEntries = iNumItems;
			auto iSelectedItem = -1;

			for (auto iArrayPos = 0; iArrayPos < iNumItems; iArrayPos++)
			{
				lpTempList->aEntries[iArrayPos].ulRowFlags = ulRowFlags;
				lpTempList->aEntries[iArrayPos].cValues = 0;
				lpTempList->aEntries[iArrayPos].rgPropVals = nullptr;
				iSelectedItem = m_lpContentsTableListCtrl->GetNextItem(iSelectedItem, LVNI_SELECTED);
				if (-1 != iSelectedItem)
				{
					// TODO: rewrite with GetSelectedItems
					const auto lpData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);
					if (lpData)
					{
						if (ulFlags & ACL_INCLUDE_ID && ulFlags & ACL_INCLUDE_OTHER)
						{
							lpTempList->aEntries[iArrayPos].rgPropVals =
								mapi::allocate<LPSPropValue>(2 * sizeof(SPropValue), lpTempList);
							lpTempList->aEntries[iArrayPos].cValues = 2;

							auto lpSPropValue =
								PpropFindProp(lpData->lpSourceProps, lpData->cSourceProps, PR_MEMBER_ID);

							lpTempList->aEntries[iArrayPos].rgPropVals[0].ulPropTag = lpSPropValue->ulPropTag;
							lpTempList->aEntries[iArrayPos].rgPropVals[0].Value = lpSPropValue->Value;

							lpSPropValue = PpropFindProp(lpData->lpSourceProps, lpData->cSourceProps, PR_MEMBER_RIGHTS);

							lpTempList->aEntries[iArrayPos].rgPropVals[1].ulPropTag = lpSPropValue->ulPropTag;
							lpTempList->aEntries[iArrayPos].rgPropVals[1].Value = lpSPropValue->Value;
						}
						else if (ulFlags & ACL_INCLUDE_ID)
						{
							lpTempList->aEntries[iArrayPos].cValues = 1;
							lpTempList->aEntries[iArrayPos].rgPropVals =
								PpropFindProp(lpData->lpSourceProps, lpData->cSourceProps, PR_MEMBER_ID);
						}
					}
				}
			}
		}

		*lppRowList = lpTempList;
		return S_OK;
	}

	void CAclDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP /*lpMAPIProp*/,
		_In_ LPMAPICONTAINER /*lpContainer*/)
	{
		if (lpParams)
		{
			lpParams->lpExchTbl = m_lpExchTbl;
		}

		ui::addinui::InvokeAddInMenu(lpParams);
	}
} // namespace dialog