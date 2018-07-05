// Displays the recipient table for a message
#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/RecipientsDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/Cache/MapiObjects.h>
#include <MAPI/ColumnTags.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <Interpret/InterpretProp.h>
#include <UI/Controls/SortList/ContentsData.h>
#include <MAPI/MAPIFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CRecipientsDlg";

	CRecipientsDlg::CRecipientsDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_ LPMAPITABLE lpMAPITable,
		_In_ LPMAPIPROP lpMessage)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_RECIPIENTS,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  lpMAPITable,
			  LPSPropTagArray(&columns::sptDEFCols),
			  columns::DEFColumns,
			  IDR_MENU_RECIPIENTS_POPUP,
			  MENU_CONTEXT_RECIPIENT_TABLE)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpMessage = mapi::safe_cast<LPMESSAGE>(lpMessage);
		m_bIsAB = true; // Recipients are from the AB
		m_bViewRecipientABEntry = false;

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_RECIPIENTS);
	}

	CRecipientsDlg::~CRecipientsDlg()
	{
		TRACE_DESTRUCTOR(CLASS);

		if (m_lpMessage)
		{
			EC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			m_lpMessage->Release();
		}
	}

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
				const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
				pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_RECIPOPTIONS, DIMMSOK(1 == iNumSel));
				pMenu->EnableMenuItem(
					ID_MODIFYRECIPIENT, DIM(1 == iNumSel && m_lpPropDisplay && m_lpPropDisplay->IsModifiedPropVals()));
			}

			pMenu->CheckMenuItem(ID_VIEWRECIPIENTABENTRY, CHECK(m_bViewRecipientABEntry));
		}

		CContentsTableDlg::OnInitMenu(pMenu);
	}

	_Check_return_ HRESULT CRecipientsDlg::OpenItemProp(
		int iSelectedItem,
		__mfcmapiModifyEnum bModify,
		_Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
	{
		output::DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		if (!m_lpContentsTableListCtrl || !lppMAPIProp) return MAPI_E_INVALID_PARAMETER;

		*lppMAPIProp = nullptr;

		if (m_bViewRecipientABEntry) return CContentsTableDlg::OpenItemProp(iSelectedItem, bModify, lppMAPIProp);

		// Do nothing, ensuring we work with the row
		return S_OK;
	}

	void CRecipientsDlg::OnViewRecipientABEntry()
	{
		m_bViewRecipientABEntry = !m_bViewRecipientABEntry;
		OnRefreshView();
	}

	void CRecipientsDlg::OnDeleteSelectedItem()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMessage || !m_lpContentsTableListCtrl) return;

		LPADRLIST lpAdrList = nullptr;

		const int iNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

		if (iNumSelected && iNumSelected < MAXNewADRLIST)
		{
			EC_H(MAPIAllocateBuffer(CbNewADRLIST(iNumSelected), reinterpret_cast<LPVOID*>(&lpAdrList)));
			if (lpAdrList)
			{
				ZeroMemory(lpAdrList, CbNewADRLIST(iNumSelected));
				lpAdrList->cEntries = iNumSelected;

				for (auto iSelection = 0; iSelection < iNumSelected; iSelection++)
				{
					LPSPropValue lpProp = nullptr;
					EC_H(MAPIAllocateBuffer(sizeof(SPropValue), reinterpret_cast<LPVOID*>(&lpProp)));

					if (lpProp)
					{
						lpAdrList->aEntries[iSelection].ulReserved1 = 0;
						lpAdrList->aEntries[iSelection].cValues = 1;
						lpAdrList->aEntries[iSelection].rgPropVals = lpProp;
						lpProp->ulPropTag = PR_ROWID;
						lpProp->dwAlignPad = 0;
						// Find the highlighted item AttachNum
						const auto lpListData = m_lpContentsTableListCtrl->GetFirstSelectedItemData();
						if (lpListData && lpListData->Contents())
						{
							lpProp->Value.l = lpListData->Contents()->m_ulRowID;
						}
						else
						{
							lpProp->Value.l = 0;
						}

						output::DebugPrintEx(
							DBGDeleteSelectedItem,
							CLASS,
							L"OnDeleteSelectedItem",
							L"Deleting row 0x%08X\n",
							lpProp->Value.l);
					}
				}

				EC_MAPI_S(m_lpMessage->ModifyRecipients(MODRECIP_REMOVE, lpAdrList));

				OnRefreshView();
				FreePadrlist(lpAdrList);
			}
		}
	}

	void CRecipientsDlg::OnModifyRecipients()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMessage || !m_lpContentsTableListCtrl || !m_lpPropDisplay) return;

		if (1 != m_lpContentsTableListCtrl->GetSelectedCount()) return;

		if (!m_lpPropDisplay->IsModifiedPropVals()) return;

		ULONG cProps = 0;
		LPSPropValue lpProps = nullptr;
		EC_H(m_lpPropDisplay->GetDisplayedProps(&cProps, &lpProps));

		if (lpProps)
		{
			ADRLIST adrList = {0};
			adrList.cEntries = 1;
			adrList.aEntries[0].ulReserved1 = 0;
			adrList.aEntries[0].cValues = cProps;

			ULONG ulSizeProps = NULL;
			hRes = EC_MAPI(ScCountProps(adrList.aEntries[0].cValues, lpProps, &ulSizeProps));

			if (SUCCEEDED(hRes))
			{
				EC_H(MAPIAllocateBuffer(ulSizeProps, reinterpret_cast<LPVOID*>(&adrList.aEntries[0].rgPropVals)));
			}

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(
					ScCopyProps(adrList.aEntries[0].cValues, lpProps, adrList.aEntries[0].rgPropVals, &ulSizeProps));
			}

			if (SUCCEEDED(hRes))
			{
				output::DebugPrintEx(
					DBGGeneric, CLASS, L"OnModifyRecipients", L"Committing changes for current selection\n");
				hRes = EC_MAPI(m_lpMessage->ModifyRecipients(MODRECIP_MODIFY, &adrList));
			}

			MAPIFreeBuffer(adrList.aEntries[0].rgPropVals);

			OnRefreshView();
		}
	}

	void CRecipientsDlg::OnRecipOptions()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMessage || !m_lpContentsTableListCtrl || !m_lpPropDisplay || !m_lpMapiObjects) return;

		if (1 != m_lpContentsTableListCtrl->GetSelectedCount()) return;

		ULONG cProps = 0;
		LPSPropValue lpProps = nullptr;
		EC_H(m_lpPropDisplay->GetDisplayedProps(&cProps, &lpProps));

		if (lpProps)
		{
			auto lpAB = m_lpMapiObjects->GetAddrBook(true); // do not release
			if (lpAB)
			{
				ADRENTRY adrEntry = {0};
				adrEntry.ulReserved1 = 0;
				adrEntry.cValues = cProps;
				adrEntry.rgPropVals = lpProps;
				output::DebugPrintEx(DBGGeneric, CLASS, L"OnRecipOptions", L"Calling RecipOptions\n");

				hRes = EC_MAPI(lpAB->RecipOptions(reinterpret_cast<ULONG_PTR>(m_hWnd), NULL, &adrEntry));

				if (hRes == MAPI_W_ERRORS_RETURNED)
				{
					LPMAPIERROR lpErr = nullptr;
					hRes = WC_MAPI(lpAB->GetLastError(hRes, fMapiUnicode, &lpErr));
					if (lpErr)
					{
						EC_MAPIERR(fMapiUnicode, lpErr);
						MAPIFreeBuffer(lpErr);
					}
					else
						CHECKHRES(hRes);
				}
				else if (SUCCEEDED(hRes))
				{
					ADRLIST adrList = {0};
					adrList.cEntries = 1;
					adrList.aEntries[0].ulReserved1 = 0;
					adrList.aEntries[0].cValues = adrEntry.cValues;
					adrList.aEntries[0].rgPropVals = adrEntry.rgPropVals;

					const auto szAdrList = interpretprop::AdrListToString(adrList);

					output::DebugPrintEx(
						DBGGeneric, CLASS, L"OnRecipOptions", L"RecipOptions returned the following ADRLIST:\n");
					// Note - debug output may be truncated due to limitations of OutputDebugString,
					// but output to file is complete
					output::Output(DBGGeneric, nullptr, false, szAdrList);

					EC_MAPI_S(m_lpMessage->ModifyRecipients(MODRECIP_MODIFY, &adrList));

					OnRefreshView();
				}
			}
		}
	}

	void CRecipientsDlg::OnSaveChanges()
	{
		if (m_lpMessage)
		{
			EC_MAPI_S(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		}
	}
}