// RecipientsDlg.cpp : implementation file
// Displays the recipient table for a message

#include "stdafx.h"
#include "Error.h"

#include "RecipientsDlg.h"

#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "ColumnTags.h"
#include "SingleMAPIPropListCtrl.h"
#include "InterpretProp.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CRecipientsDlg");

/////////////////////////////////////////////////////////////////////////////
// CRecipientsDlg dialog

CRecipientsDlg::CRecipientsDlg(
							   CParentWnd* pParentWnd,
							   CMapiObjects *lpMapiObjects,
							   LPMAPITABLE lpMAPITable,
							   LPMESSAGE lpMessage
							   ):
CContentsTableDlg(
						  pParentWnd,
						  lpMapiObjects,
						  IDS_RECIPIENTS,
						  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
						  lpMAPITable,
						  (LPSPropTagArray) &sptDEFCols,
						  NUMDEFCOLUMNS,
						  DEFColumns,
						  IDR_MENU_RECIPIENTS_POPUP,
						  MENU_CONTEXT_RECIPIENT_TABLE)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpMessage = lpMessage;
	if (m_lpMessage) m_lpMessage->AddRef();
	m_bIsAB = true;//Recipients are from the AB

	CreateDialogAndMenu(IDR_MENU_RECIPIENTS);
}

CRecipientsDlg::~CRecipientsDlg()
{
	TRACE_DESTRUCTOR(CLASS);

//	if (m_lpMessage)
//	{
//		HRESULT hRes = S_OK;
//
//		EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
//	}
	if (m_lpMessage) m_lpMessage->Release();
}

BEGIN_MESSAGE_MAP(CRecipientsDlg, CContentsTableDlg)
//{{AFX_MSG_MAP(CRecipientsDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYRECIPIENT, OnModifyRecipients)
	ON_COMMAND(ID_RECIPOPTIONS, OnRecipOptions)
	ON_COMMAND(ID_SAVECHANGES, OnSaveChanges)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void CRecipientsDlg::OnInitMenu(CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_MODIFYRECIPIENT,
				DIM(1 == iNumSel && m_lpPropDisplay && m_lpPropDisplay->IsModifiedPropVals()));
		}
	}
	CContentsTableDlg::OnInitMenu(pMenu);
}

void CRecipientsDlg::OnDeleteSelectedItem()
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	if (!m_lpMessage || !m_lpContentsTableListCtrl) return;

	int				iItem = -1;
	SortListData*	lpListData = NULL;
	LPADRLIST		lpAdrList = NULL;

	int iNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

	if (iNumSelected && iNumSelected < MAXNewADRLIST)
	{
		EC_H(MAPIAllocateBuffer(
			(ULONG)CbNewADRLIST(iNumSelected),
			(LPVOID*) &lpAdrList));
		if (lpAdrList)
		{
			ZeroMemory(lpAdrList, CbNewADRLIST(iNumSelected));
			lpAdrList->cEntries = iNumSelected;

			LPSPropValue lpProps = NULL;
			EC_H(MAPIAllocateBuffer(
				(ULONG)(iNumSelected * sizeof(SPropValue)),
				(LPVOID*) &lpProps));

			if (lpProps)
			{
				int iSelection = 0;
				for (iSelection = 0 ; iSelection < iNumSelected ; iSelection++)
				{
					lpAdrList->aEntries[iSelection].ulReserved1 = 0;
					lpAdrList->aEntries[iSelection].cValues = 1;
					lpAdrList->aEntries[iSelection].rgPropVals = &lpProps[iSelection];
					lpProps[iSelection].ulPropTag = PR_ROWID;
					lpProps[iSelection].dwAlignPad = 0;
					//Find the highlighted item AttachNum
					lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
					if (lpListData)
					{
						lpProps[iSelection].Value.l = lpListData->data.Contents.ulRowID;
					}
					else
					{
						lpProps[iSelection].Value.l = 0;
					}
					DebugPrintEx(DBGDeleteSelectedItem,CLASS,_T("OnDeleteSelectedItem"),_T("Deleting row 0x%08X\n"),lpProps[iSelection].Value.l);
				}

				EC_H(m_lpMessage->ModifyRecipients(
					MODRECIP_REMOVE,
					lpAdrList));

				//EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			}
			FreePadrlist(lpAdrList);
		}
	}
}//CRecipientsDlg::OnDeleteSelectedItem

void CRecipientsDlg::OnModifyRecipients()
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	if (!m_lpMessage || !m_lpContentsTableListCtrl || !m_lpPropDisplay) return;

	if (1 != m_lpContentsTableListCtrl->GetSelectedCount()) return;

	if (!m_lpPropDisplay->IsModifiedPropVals()) return;

	LPSPropValue lpProps = m_lpPropDisplay->GetPropVals();

	if (lpProps)
	{
		ADRLIST	adrList = {0};
		adrList.cEntries = 1;
		adrList.aEntries[0].ulReserved1 = 0;
		adrList.aEntries[0].cValues = m_lpPropDisplay->GetCountPropVals();
		adrList.aEntries[0].rgPropVals = lpProps;
		DebugPrintEx(DBGGeneric,CLASS,_T("OnModifyRecipients"),_T("Committing changes for current selection\n"));

		EC_H(m_lpMessage->ModifyRecipients(
			MODRECIP_MODIFY,
			&adrList));

//		EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}
}

void CRecipientsDlg::OnRecipOptions()
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	if (!m_lpMessage || !m_lpContentsTableListCtrl || !m_lpPropDisplay || !m_lpMapiObjects) return;

	if (1 != m_lpContentsTableListCtrl->GetSelectedCount()) return;

	LPSPropValue lpProps = m_lpPropDisplay->GetPropVals();

	if (lpProps)
	{
		LPADRBOOK lpAB = m_lpMapiObjects->GetAddrBook(true);//do not release
		if (lpAB)
		{
			ADRENTRY adrEntry = {0};
			adrEntry.ulReserved1 = 0;
			adrEntry.cValues = m_lpPropDisplay->GetCountPropVals();
			adrEntry.rgPropVals = lpProps;
			DebugPrintEx(DBGGeneric,CLASS,_T("OnRecipOptions"),_T("Calling RecipOptions\n"));

			EC_H(lpAB->RecipOptions(
				(ULONG_PTR) m_hWnd,
				NULL,
				&adrEntry));

			if (MAPI_W_ERRORS_RETURNED == hRes)
			{
				LPMAPIERROR lpErr = NULL;
				hRes = lpAB->GetLastError(hRes,fMapiUnicode,&lpErr);
				if (lpErr)
				{
					EC_MAPIERR(fMapiUnicode,lpErr);
					MAPIFreeBuffer(lpErr);
				}
				else CHECKHRES(hRes);
			}
			else if (SUCCEEDED(hRes))
			{
				ADRLIST adrList = {0};
				adrList.cEntries = 1;
				adrList.aEntries[0].ulReserved1 = 0;
				adrList.aEntries[0].cValues = adrEntry.cValues;
				adrList.aEntries[0].rgPropVals = adrEntry.rgPropVals;

				CString szAdrList;
				AdrListToString(&adrList,&szAdrList);

				DebugPrintEx(DBGGeneric,CLASS,_T("OnRecipOptions"),_T("RecipOptions returned the following ADRLIST:\n"));
				// This buffer may be huge - passing it as single parameter to _Output avoids calls to StringCchVPrintf
				// Note - debug output may still be truncated due to limitations of OutputDebugString,
				// but output to file is complete
				_Output(DBGGeneric,NULL,false,(LPCTSTR)szAdrList);

				EC_H(m_lpMessage->ModifyRecipients(
					MODRECIP_MODIFY,
					&adrList));

				OnRefreshView();
			}
		}
	}
}

void CRecipientsDlg::OnSaveChanges()
{
	if (m_lpMessage)
	{
		HRESULT hRes = S_OK;

		EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}
}

