// Displays the rules table for a folder
#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/RulesDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/Cache/MapiObjects.h>
#include <MAPI/ColumnTags.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <MAPI/MAPIFunctions.h>
#include <MAPI/MapiMemory.h>
#include <UI/addinui.h>

namespace dialog
{
	static std::wstring CLASS = L"CRulesDlg";

#define RULE_INCLUDE_ID 0x00000001
#define RULE_INCLUDE_OTHER 0x00000002

	CRulesDlg::CRulesDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_ LPEXCHANGEMODIFYTABLE lpExchTbl)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_RULESTABLE,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  nullptr,
			  &columns::sptRULECols.tags,
			  columns::RULEColumns,
			  IDR_MENU_RULES_POPUP,
			  MENU_CONTEXT_RULES_TABLE)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpExchTbl = lpExchTbl;
		if (m_lpExchTbl) m_lpExchTbl->AddRef();
		m_bIsAB = false;

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_RULES);

		CRulesDlg::OnRefreshView();
	}

	CRulesDlg::~CRulesDlg()
	{
		TRACE_DESTRUCTOR(CLASS);

		if (m_lpExchTbl) m_lpExchTbl->Release();
	}

	BEGIN_MESSAGE_MAP(CRulesDlg, CContentsTableDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYSELECTEDITEM, OnModifySelectedItem)
	END_MESSAGE_MAP()

	void CRulesDlg::OnInitMenu(_In_opt_ CMenu* pMenu)
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
	void CRulesDlg::OnRefreshView()
	{
		if (!m_lpExchTbl || !m_lpContentsTableListCtrl) return;

		if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L"\n");

		if (m_lpExchTbl)
		{
			LPMAPITABLE lpMAPITable = nullptr;
			// Open a MAPI table on the Exchange table property. This table can be
			// read to determine what the Exchange table looks like.
			EC_MAPI_S(m_lpExchTbl->GetTable(0, &lpMAPITable));
			if (lpMAPITable)
			{
				m_lpContentsTableListCtrl->SetContentsTable(lpMAPITable, dfDeleted, NULL);

				lpMAPITable->Release();
			}
		}
	}

	void CRulesDlg::OnDeleteSelectedItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		LPROWLIST lpSelectedItems = nullptr;

		EC_H_S(GetSelectedItems(RULE_INCLUDE_ID, ROW_REMOVE, &lpSelectedItems));
		if (lpSelectedItems)
		{
			const auto hRes = EC_MAPI(m_lpExchTbl->ModifyTable(0, lpSelectedItems));
			MAPIFreeBuffer(lpSelectedItems);
			if (hRes == S_OK) OnRefreshView();
		}
	}

	void CRulesDlg::OnModifySelectedItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		LPROWLIST lpSelectedItems = nullptr;

		EC_H_S(GetSelectedItems(RULE_INCLUDE_ID | RULE_INCLUDE_OTHER, ROW_MODIFY, &lpSelectedItems));
		if (lpSelectedItems)
		{
			const auto hRes = EC_MAPI(m_lpExchTbl->ModifyTable(0, lpSelectedItems));
			MAPIFreeBuffer(lpSelectedItems);
			if (hRes == S_OK) OnRefreshView();
		}
	}

	_Check_return_ HRESULT
	CRulesDlg::GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, _In_ LPROWLIST* lppRowList) const
	{
		if (!lppRowList || !m_lpContentsTableListCtrl) return MAPI_E_INVALID_PARAMETER;
		*lppRowList = nullptr;
		const int iNumItems = m_lpContentsTableListCtrl->GetSelectedCount();

		if (!iNumItems) return S_OK;
		if (iNumItems > MAXNewROWLIST) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
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
					// TODO: Rewrite with GetSelectedItems
					const auto lpData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);
					if (lpData)
					{
						if (ulFlags & RULE_INCLUDE_ID && ulFlags & RULE_INCLUDE_OTHER)
						{
							lpTempList->aEntries[iArrayPos].rgPropVals =
								mapi::allocate<LPSPropValue>(lpData->cSourceProps * sizeof(SPropValue), lpTempList);
							if (lpTempList->aEntries[iArrayPos].rgPropVals)
							{
								ULONG ulDst = 0;
								for (ULONG ulSrc = 0; ulSrc < lpData->cSourceProps; ulSrc++)
								{
									if (lpData->lpSourceProps[ulSrc].ulPropTag == PR_RULE_PROVIDER_DATA)
									{
										if (!lpData->lpSourceProps[ulSrc].Value.bin.cb ||
											!lpData->lpSourceProps[ulSrc].Value.bin.lpb)
										{
											// PR_RULE_PROVIDER_DATA was NULL - we don't want this
											continue;
										}
									}

									hRes = EC_H(mapi::MyPropCopyMore(
										&lpTempList->aEntries[iArrayPos].rgPropVals[ulDst],
										&lpData->lpSourceProps[ulSrc],
										MAPIAllocateMore,
										lpTempList));
									ulDst++;
								}

								lpTempList->aEntries[iArrayPos].cValues = ulDst;
							}
						}
						else if (ulFlags & RULE_INCLUDE_ID)
						{
							lpTempList->aEntries[iArrayPos].cValues = 1;
							lpTempList->aEntries[iArrayPos].rgPropVals =
								PpropFindProp(lpData->lpSourceProps, lpData->cSourceProps, PR_RULE_ID);
						}
					}
				}
			}
		}

		*lppRowList = lpTempList;
		return hRes;
	}

	void CRulesDlg::HandleAddInMenuSingle(
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
}
