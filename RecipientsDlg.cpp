// RecipientsDlg.cpp : implementation file
// Displays the recipient table for a message

#include "stdafx.h"
#include "RecipientsDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "ColumnTags.h"
#include "SingleMAPIPropListCtrl.h"
#include "InterpretProp.h"
#include "string.h"

static wstring CLASS = L"CRecipientsDlg";

/////////////////////////////////////////////////////////////////////////////
// CRecipientsDlg dialog

CRecipientsDlg::CRecipientsDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPMAPITABLE lpMAPITable,
	_In_ LPMESSAGE lpMessage
	) :
	CContentsTableDlg(
	pParentWnd,
	lpMapiObjects,
	IDS_RECIPIENTS,
	mfcmapiDO_NOT_CALL_CREATE_DIALOG,
	lpMAPITable,
	(LPSPropTagArray)&sptDEFCols,
	NUMDEFCOLUMNS,
	DEFColumns,
	IDR_MENU_RECIPIENTS_POPUP,
	MENU_CONTEXT_RECIPIENT_TABLE)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpMessage = lpMessage;
	if (m_lpMessage) m_lpMessage->AddRef();
	m_bIsAB = true; // Recipients are from the AB
	m_bViewRecipientABEntry = false;

	CreateDialogAndMenu(IDR_MENU_RECIPIENTS);
} // CRecipientsDlg::CRecipientsDlg

CRecipientsDlg::~CRecipientsDlg()
{
	TRACE_DESTRUCTOR(CLASS);

	if (m_lpMessage)
	{
		HRESULT hRes = S_OK;

		EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		m_lpMessage->Release();
	}
} // CRecipientsDlg::~CRecipientsDlg

BEGIN_MESSAGE_MAP(CRecipientsDlg, CContentsTableDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYRECIPIENT, OnModifyRecipients)
	ON_COMMAND(ID_RECIPOPTIONS, OnRecipOptions)
	ON_COMMAND(ID_SAVECHANGES, OnSaveChanges)
	ON_COMMAND(ID_VIEWRECIPIENTABENTRY, OnViewRecipientABEntry)
END_MESSAGE_MAP()

void CRecipientsDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_RECIPOPTIONS, DIMMSOK(1 == iNumSel));
			pMenu->EnableMenuItem(ID_MODIFYRECIPIENT,
				DIM(1 == iNumSel && m_lpPropDisplay && m_lpPropDisplay->IsModifiedPropVals()));
		}
		pMenu->CheckMenuItem(ID_VIEWRECIPIENTABENTRY, CHECK(m_bViewRecipientABEntry));
	}
	CContentsTableDlg::OnInitMenu(pMenu);
} // CRecipientsDlg::OnInitMenu

_Check_return_ HRESULT CRecipientsDlg::OpenItemProp(
	int iSelectedItem,
	__mfcmapiModifyEnum bModify,
	_Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

	if (!m_lpContentsTableListCtrl || !lppMAPIProp) return MAPI_E_INVALID_PARAMETER;

	*lppMAPIProp = NULL;

	if (m_bViewRecipientABEntry) return CContentsTableDlg::OpenItemProp(iSelectedItem, bModify, lppMAPIProp);

	// Do nothing, ensuring we work with the row
	return S_OK;
} // CRecipientsDlg::OpenItemProp

void CRecipientsDlg::OnViewRecipientABEntry()
{
	m_bViewRecipientABEntry = !m_bViewRecipientABEntry;
	OnRefreshView();
} // CRecipientsDlg::OnViewRecipientABEntry

void CRecipientsDlg::OnDeleteSelectedItem()
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMessage || !m_lpContentsTableListCtrl) return;

	int				iItem = -1;
	SortListData*	lpListData = NULL;
	LPADRLIST		lpAdrList = NULL;

	int iNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

	if (iNumSelected && iNumSelected < MAXNewADRLIST)
	{
		EC_H(MAPIAllocateBuffer(
			(ULONG)CbNewADRLIST(iNumSelected),
			(LPVOID*)&lpAdrList));
		if (lpAdrList)
		{
			ZeroMemory(lpAdrList, CbNewADRLIST(iNumSelected));
			lpAdrList->cEntries = iNumSelected;

			int iSelection = 0;
			for (iSelection = 0; iSelection < iNumSelected; iSelection++)
			{
				LPSPropValue lpProp = NULL;
				EC_H(MAPIAllocateBuffer(
					sizeof(SPropValue),
					(LPVOID*)&lpProp));

				if (lpProp)
				{
					lpAdrList->aEntries[iSelection].ulReserved1 = 0;
					lpAdrList->aEntries[iSelection].cValues = 1;
					lpAdrList->aEntries[iSelection].rgPropVals = lpProp;
					lpProp->ulPropTag = PR_ROWID;
					lpProp->dwAlignPad = 0;
					// Find the highlighted item AttachNum
					lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
					if (lpListData)
					{
						lpProp->Value.l = lpListData->data.Contents.ulRowID;
					}
					else
					{
						lpProp->Value.l = 0;
					}
					DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L"Deleting row 0x%08X\n", lpProp->Value.l);
				}
			}

			EC_MAPI(m_lpMessage->ModifyRecipients(
				MODRECIP_REMOVE,
				lpAdrList));

			OnRefreshView();
			FreePadrlist(lpAdrList);
		}
	}
} // CRecipientsDlg::OnDeleteSelectedItem

void CRecipientsDlg::OnModifyRecipients()
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMessage || !m_lpContentsTableListCtrl || !m_lpPropDisplay) return;

	if (1 != m_lpContentsTableListCtrl->GetSelectedCount()) return;

	if (!m_lpPropDisplay->IsModifiedPropVals()) return;

	ULONG cProps = 0;
	LPSPropValue lpProps = NULL;
	EC_H(m_lpPropDisplay->GetDisplayedProps(&cProps, &lpProps));

	if (lpProps)
	{
		ADRLIST	adrList = { 0 };
		adrList.cEntries = 1;
		adrList.aEntries[0].ulReserved1 = 0;
		adrList.aEntries[0].cValues = cProps;

		ULONG ulSizeProps = NULL;
		EC_MAPI(ScCountProps(
			adrList.aEntries[0].cValues,
			lpProps,
			&ulSizeProps));

		EC_H(MAPIAllocateBuffer(ulSizeProps, (LPVOID*)&adrList.aEntries[0].rgPropVals));

		EC_MAPI(ScCopyProps(
			adrList.aEntries[0].cValues,
			lpProps,
			adrList.aEntries[0].rgPropVals,
			&ulSizeProps));

		DebugPrintEx(DBGGeneric, CLASS, L"OnModifyRecipients", L"Committing changes for current selection\n");

		EC_MAPI(m_lpMessage->ModifyRecipients(
			MODRECIP_MODIFY,
			&adrList));

		MAPIFreeBuffer(adrList.aEntries[0].rgPropVals);

		OnRefreshView();
	}
} // CRecipientsDlg::OnModifyRecipients

void CRecipientsDlg::OnRecipOptions()
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMessage || !m_lpContentsTableListCtrl || !m_lpPropDisplay || !m_lpMapiObjects) return;

	if (1 != m_lpContentsTableListCtrl->GetSelectedCount()) return;

	ULONG cProps = 0;
	LPSPropValue lpProps = NULL;
	EC_H(m_lpPropDisplay->GetDisplayedProps(&cProps, &lpProps));

	if (lpProps)
	{
		LPADRBOOK lpAB = m_lpMapiObjects->GetAddrBook(true); // do not release
		if (lpAB)
		{
			ADRENTRY adrEntry = { 0 };
			adrEntry.ulReserved1 = 0;
			adrEntry.cValues = cProps;
			adrEntry.rgPropVals = lpProps;
			DebugPrintEx(DBGGeneric, CLASS, L"OnRecipOptions", L"Calling RecipOptions\n");

			EC_MAPI(lpAB->RecipOptions(
				(ULONG_PTR)m_hWnd,
				NULL,
				&adrEntry));

			if (MAPI_W_ERRORS_RETURNED == hRes)
			{
				LPMAPIERROR lpErr = NULL;
				WC_MAPI(lpAB->GetLastError(hRes, fMapiUnicode, &lpErr));
				if (lpErr)
				{
					EC_MAPIERR(fMapiUnicode, lpErr);
					MAPIFreeBuffer(lpErr);
				}
				else CHECKHRES(hRes);
			}
			else if (SUCCEEDED(hRes))
			{
				ADRLIST adrList = { 0 };
				adrList.cEntries = 1;
				adrList.aEntries[0].ulReserved1 = 0;
				adrList.aEntries[0].cValues = adrEntry.cValues;
				adrList.aEntries[0].rgPropVals = adrEntry.rgPropVals;

				wstring szAdrList = AdrListToString(&adrList);

				DebugPrintEx(DBGGeneric, CLASS, L"OnRecipOptions", L"RecipOptions returned the following ADRLIST:\n");
				// Note - debug output may be truncated due to limitations of OutputDebugString,
				// but output to file is complete
				_OutputW(DBGGeneric, NULL, false, szAdrList);

				EC_MAPI(m_lpMessage->ModifyRecipients(
					MODRECIP_MODIFY,
					&adrList));

				OnRefreshView();
			}
		}
	}
} // CRecipientsDlg::OnRecipOptions

void CRecipientsDlg::OnSaveChanges()
{
	if (m_lpMessage)
	{
		HRESULT hRes = S_OK;

		EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}
} // CRecipientsDlg::OnSaveChanges