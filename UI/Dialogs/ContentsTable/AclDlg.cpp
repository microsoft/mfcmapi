// Displays the ACL table for an item

#include "StdAfx.h"
#include "AclDlg.h"
#include <UI/Controls/ContentsTableListCtrl.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <MAPI/MapiObjects.h>
#include <MAPI/ColumnTags.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <Interpret/InterpretProp.h>

static std::wstring CLASS = L"CAclDlg";

#define ACL_INCLUDE_ID 0x00000001
#define ACL_INCLUDE_OTHER 0x00000002

CAclDlg::CAclDlg(_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPEXCHANGEMODIFYTABLE lpExchTbl,
	bool fFreeBusyVisible)
	: CContentsTableDlg(pParentWnd,
		lpMapiObjects,
		fFreeBusyVisible ? IDS_ACLFBTABLE : IDS_ACLTABLE,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		nullptr,
		LPSPropTagArray(&sptACLCols),
		ACLColumns,
		IDR_MENU_ACL_POPUP,
		MENU_CONTEXT_ACL_TABLE)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpExchTbl = lpExchTbl;

	if (m_lpExchTbl)
		m_lpExchTbl->AddRef();

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

	if (m_lpExchTbl)
		m_lpExchTbl->Release();
}

BEGIN_MESSAGE_MAP(CAclDlg, CContentsTableDlg)
	ON_COMMAND(ID_ADDITEM, OnAddItem)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYSELECTEDITEM, OnModifySelectedItem)
END_MESSAGE_MAP()

_Check_return_ HRESULT CAclDlg::OpenItemProp(int /*iSelectedItem*/, __mfcmapiModifyEnum /*bModify*/, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	if (lppMAPIProp) *lppMAPIProp = nullptr;
	// Don't do anything because we don't want to override the properties that we have
	return S_OK;
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
	auto hRes = S_OK;

	if (!m_lpExchTbl || !m_lpContentsTableListCtrl)
		return;

	if (m_lpContentsTableListCtrl->IsLoading())
		m_lpContentsTableListCtrl->OnCancelTableLoad();
	DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L"\n");

	if (m_lpExchTbl)
	{
		LPMAPITABLE lpMAPITable = nullptr;
		// Open a MAPI table on the Exchange table property. This table can be
		// read to determine what the Exchange table looks like.
		EC_MAPI(m_lpExchTbl->GetTable(m_ulTableFlags, &lpMAPITable));

		if (lpMAPITable)
		{
			EC_H(m_lpContentsTableListCtrl->SetContentsTable(
				lpMAPITable,
				dfNormal,
				NULL));

			lpMAPITable->Release();
		}
	}
}

void CAclDlg::OnAddItem()
{
	auto hRes = S_OK;

	CEditor MyData(
		this,
		IDS_ACLADDITEM,
		IDS_ACLADDITEMPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.SetPromptPostFix(interpretprop::AllFlagsToString(PROP_ID(PR_MEMBER_RIGHTS), true));
	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_USEREID, false));
	MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_MASKINHEX, false));
	MyData.SetHex(1, 0);

	WC_H(MyData.DisplayDialog());
	if (S_OK != hRes)
	{
		DebugPrint(DBGGeneric, L"OnAddItem cancelled.\n");
		return;
	}

	LPROWLIST lpNewItem = nullptr;

	EC_H(MAPIAllocateBuffer(CbNewROWLIST(1), reinterpret_cast<LPVOID*>(&lpNewItem)));

	if (lpNewItem)
	{
		lpNewItem->cEntries = 1;
		lpNewItem->aEntries[0].ulRowFlags = ROW_ADD;
		lpNewItem->aEntries[0].cValues = 2;
		lpNewItem->aEntries[0].rgPropVals = nullptr;

		EC_H(MAPIAllocateMore(2 * sizeof(SPropValue), lpNewItem, reinterpret_cast<LPVOID*>(&lpNewItem->aEntries[0].rgPropVals)));

		if (lpNewItem->aEntries[0].rgPropVals)
		{
			LPENTRYID lpEntryID = nullptr;
			size_t cbBin = 0;
			EC_H(MyData.GetEntryID(0, false, &cbBin, &lpEntryID));

			lpNewItem->aEntries[0].rgPropVals[0].ulPropTag = PR_MEMBER_ENTRYID;
			lpNewItem->aEntries[0].rgPropVals[0].Value.bin.cb = static_cast<ULONG>(cbBin);
			lpNewItem->aEntries[0].rgPropVals[0].Value.bin.lpb = reinterpret_cast<LPBYTE>(lpEntryID);
			lpNewItem->aEntries[0].rgPropVals[1].ulPropTag = PR_MEMBER_RIGHTS;
			lpNewItem->aEntries[0].rgPropVals[1].Value.ul = MyData.GetHex(1);

			EC_MAPI(m_lpExchTbl->ModifyTable(
				m_ulTableFlags,
				lpNewItem));
			MAPIFreeBuffer(lpNewItem);
			if (S_OK == hRes)
				OnRefreshView();

			delete[] lpEntryID;
		}
	}
}

void CAclDlg::OnDeleteSelectedItem()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	LPROWLIST lpSelectedItems = nullptr;

	EC_H(GetSelectedItems(ACL_INCLUDE_ID, ROW_REMOVE, &lpSelectedItems));

	if (lpSelectedItems)
	{
		EC_MAPI(m_lpExchTbl->ModifyTable(
			m_ulTableFlags,
			lpSelectedItems));
		MAPIFreeBuffer(lpSelectedItems);
		if (S_OK == hRes)
			OnRefreshView();
	}
}


void CAclDlg::OnModifySelectedItem()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	LPROWLIST lpSelectedItems = nullptr;

	EC_H(GetSelectedItems(ACL_INCLUDE_ID | ACL_INCLUDE_OTHER, ROW_MODIFY, &lpSelectedItems));

	if (lpSelectedItems)
	{
		EC_MAPI(m_lpExchTbl->ModifyTable(
			m_ulTableFlags,
			lpSelectedItems));
		MAPIFreeBuffer(lpSelectedItems);
		if (S_OK == hRes) OnRefreshView();
	}
}

_Check_return_ HRESULT CAclDlg::GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, _In_ LPROWLIST* lppRowList) const
{
	if (!lppRowList || !m_lpContentsTableListCtrl)
		return MAPI_E_INVALID_PARAMETER;

	*lppRowList = nullptr;
	auto hRes = S_OK;
	const int iNumItems = m_lpContentsTableListCtrl->GetSelectedCount();

	if (!iNumItems) return S_OK;
	if (iNumItems > MAXNewROWLIST) return MAPI_E_INVALID_PARAMETER;

	LPROWLIST lpTempList = nullptr;

	EC_H(MAPIAllocateBuffer(CbNewROWLIST(iNumItems), reinterpret_cast<LPVOID*>(&lpTempList)));

	if (lpTempList)
	{
		lpTempList->cEntries = iNumItems;
		auto iSelectedItem = -1;

		for (auto iArrayPos = 0; iArrayPos < iNumItems; iArrayPos++)
		{
			lpTempList->aEntries[iArrayPos].ulRowFlags = ulRowFlags;
			lpTempList->aEntries[iArrayPos].cValues = 0;
			lpTempList->aEntries[iArrayPos].rgPropVals = nullptr;
			iSelectedItem = m_lpContentsTableListCtrl->GetNextItem(
				iSelectedItem,
				LVNI_SELECTED);
			if (-1 != iSelectedItem)
			{
				// TODO: rewrite with GetSelectedItems
				const auto lpData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);
				if (lpData)
				{
					if (ulFlags & ACL_INCLUDE_ID && ulFlags & ACL_INCLUDE_OTHER)
					{
						EC_H(MAPIAllocateMore(2 * sizeof(SPropValue), lpTempList, reinterpret_cast<LPVOID*>(&lpTempList->aEntries[iArrayPos].rgPropVals)));

						lpTempList->aEntries[iArrayPos].cValues = 2;

						auto lpSPropValue = PpropFindProp(
							lpData->lpSourceProps,
							lpData->cSourceProps,
							PR_MEMBER_ID);

						lpTempList->aEntries[iArrayPos].rgPropVals[0].ulPropTag = lpSPropValue->ulPropTag;
						lpTempList->aEntries[iArrayPos].rgPropVals[0].Value = lpSPropValue->Value;

						lpSPropValue = PpropFindProp(
							lpData->lpSourceProps,
							lpData->cSourceProps,
							PR_MEMBER_RIGHTS);

						lpTempList->aEntries[iArrayPos].rgPropVals[1].ulPropTag = lpSPropValue->ulPropTag;
						lpTempList->aEntries[iArrayPos].rgPropVals[1].Value = lpSPropValue->Value;
					}
					else if (ulFlags & ACL_INCLUDE_ID)
					{
						lpTempList->aEntries[iArrayPos].cValues = 1;
						lpTempList->aEntries[iArrayPos].rgPropVals = PpropFindProp(
							lpData->lpSourceProps,
							lpData->cSourceProps,
							PR_MEMBER_ID);
					}
				}
			}
		}
	}

	*lppRowList = lpTempList;
	return hRes;
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

	InvokeAddInMenu(lpParams);
}