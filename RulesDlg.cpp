// RulesDlg.cpp : implementation file
// Displays the rules table for a folder

#include "stdafx.h"
#include "Error.h"

#include "RulesDlg.h"

#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "ColumnTags.h"
#include "SingleMAPIPropListCtrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CRulesDlg");

#define RULE_INCLUDE_ID			0x00000001
#define RULE_INCLUDE_OTHER		0x00000002

/////////////////////////////////////////////////////////////////////////////
// CRulesDlg dialog

CRulesDlg::CRulesDlg(
							   CParentWnd* pParentWnd,
							   CMapiObjects *lpMapiObjects,
							   LPEXCHANGEMODIFYTABLE lpExchTbl):
CContentsTableDlg(
						  pParentWnd,
						  lpMapiObjects,
						  IDS_RULESTABLE,
						  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
						  NULL,
						  (LPSPropTagArray) &sptRULECols,
						  NUMRULECOLUMNS,
						  RULEColumns,
						  IDR_MENU_RULES_POPUP,
						  MENU_CONTEXT_RULES_TABLE)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpExchTbl = lpExchTbl;
	if (m_lpExchTbl) m_lpExchTbl->AddRef();
	m_bIsAB = false;

	CreateDialogAndMenu(IDR_MENU_RULES);

	OnRefreshView();
}

CRulesDlg::~CRulesDlg()
{
	TRACE_DESTRUCTOR(CLASS);

	if (m_lpExchTbl) m_lpExchTbl->Release();
}

BEGIN_MESSAGE_MAP(CRulesDlg, CContentsTableDlg)
//{{AFX_MSG_MAP(CRulesDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYSELECTEDITEM, OnModifySelectedItem)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void CRulesDlg::OnInitMenu(CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_MODIFYSELECTEDITEM,DIMMSOK(iNumSel));
		}
	}
	CContentsTableDlg::OnInitMenu(pMenu);
}


//Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void CRulesDlg::OnRefreshView()
{
	HRESULT hRes = S_OK;

	if (!m_lpExchTbl || !m_lpContentsTableListCtrl) return;

	if (m_lpContentsTableListCtrl->m_bInLoadOp) m_lpContentsTableListCtrl->OnCancelTableLoad();
	DebugPrintEx(DBGGeneric,CLASS,_T("OnRefreshView"),_T("\n"));

	if (m_lpExchTbl)
	{
		LPMAPITABLE lpMAPITable = NULL;
		// Open a MAPI table on the Exchange table property. This table can be
		// read to determine what the Exchange table looks like.
		EC_H(m_lpExchTbl->GetTable(0, &lpMAPITable));

		if (lpMAPITable)
		{
			EC_H(m_lpContentsTableListCtrl->SetContentsTable(
				lpMAPITable,
				dfDeleted,
				NULL));

			lpMAPITable->Release();
		}
	}
}//CRulesDlg::OnRefreshView

void CRulesDlg::OnDeleteSelectedItem()
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	LPROWLIST lpSelectedItems = NULL;

	EC_H(GetSelectedItems(RULE_INCLUDE_ID, ROW_REMOVE, &lpSelectedItems));

	if (lpSelectedItems)
	{
		EC_H(m_lpExchTbl->ModifyTable(
			0,
			lpSelectedItems));
		MAPIFreeBuffer(lpSelectedItems);
		if (S_OK == hRes) OnRefreshView();
	}
}//CRulesDlg::OnDeleteSelectedItem

void CRulesDlg::OnModifySelectedItem()
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	LPROWLIST lpSelectedItems = NULL;

	EC_H(GetSelectedItems(RULE_INCLUDE_ID|RULE_INCLUDE_OTHER, ROW_MODIFY, &lpSelectedItems));

	if (lpSelectedItems)
	{
		EC_H(m_lpExchTbl->ModifyTable(
			0,
			lpSelectedItems));
		MAPIFreeBuffer(lpSelectedItems);
		if (S_OK == hRes) OnRefreshView();
	}
}

HRESULT CRulesDlg::GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, LPROWLIST* lppRowList)
{
	if (!lppRowList || !m_lpContentsTableListCtrl) return MAPI_E_INVALID_PARAMETER;
	*lppRowList = NULL;
	HRESULT hRes = S_OK;
	int iNumItems = m_lpContentsTableListCtrl->GetSelectedCount();

	if (!iNumItems) return S_OK;
	if (iNumItems > MAXNewROWLIST) return MAPI_E_INVALID_PARAMETER;

	LPROWLIST lpTempList = NULL;

	EC_H(MAPIAllocateBuffer(CbNewROWLIST(iNumItems),(LPVOID*) &lpTempList));

	if (lpTempList)
	{
		lpTempList->cEntries = iNumItems;
		int iArrayPos = 0;
		int iSelectedItem = -1;

		for(iArrayPos = 0 ; iArrayPos < iNumItems ; iArrayPos++)
		{
			lpTempList->aEntries[iArrayPos].ulRowFlags = ulRowFlags;
			lpTempList->aEntries[iArrayPos].cValues = 0;
			lpTempList->aEntries[iArrayPos].rgPropVals = 0;
			iSelectedItem = m_lpContentsTableListCtrl->GetNextItem(
				iSelectedItem,
				LVNI_SELECTED);
			if (-1 != iSelectedItem)
			{
				SortListData* lpData = (SortListData*) m_lpContentsTableListCtrl->GetItemData(iSelectedItem);
				if (lpData)
				{
					if (ulFlags & RULE_INCLUDE_ID && ulFlags & RULE_INCLUDE_OTHER)
					{
						EC_H(MAPIAllocateMore(
							lpData->cSourceProps * sizeof(SPropValue),
							lpTempList,
							(LPVOID*) &lpTempList->aEntries[iArrayPos].rgPropVals));
						if (SUCCEEDED(hRes) && lpTempList->aEntries[iArrayPos].rgPropVals)
						{
							ULONG ulSrc = 0;
							ULONG ulDst = 0;
							for (ulSrc = 0; ulSrc < lpData->cSourceProps; ulSrc++)
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

								// This relies on our augmented PropCopyMore that understands PT_SRESTRICTION and PT_ACTIONS
								EC_H(PropCopyMore(
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
						lpTempList->aEntries[iArrayPos].rgPropVals = PpropFindProp(
							lpData->lpSourceProps,
							lpData->cSourceProps,
							PR_RULE_ID);
					}
				}
			}
		}
	}

	*lppRowList = lpTempList;
	return hRes;
}//CRulesDlg::GetSelectedItems

void CRulesDlg::HandleAddInMenuSingle(
									   LPADDINMENUPARAMS lpParams,
									   LPMAPIPROP /*lpMAPIProp*/,
									   LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpExchTbl = m_lpExchTbl;
	}

	InvokeAddInMenu(lpParams);
} // CRulesDlg::HandleAddInMenuSingle