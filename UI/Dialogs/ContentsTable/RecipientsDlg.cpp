// Displays the recipient table for a message
#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/RecipientsDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/columnTags.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <core/sortlistdata/contentsData.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>
#include <core/property/parseProperty.h>

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
			  &columns::sptDEFCols.tags,
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

	_Check_return_ LPMAPIPROP CRecipientsDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify)
	{
		if (!m_lpContentsTableListCtrl) return nullptr;
		output::DebugPrintEx(output::DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		if (m_bViewRecipientABEntry)
		{
			return CContentsTableDlg::OpenItemProp(iSelectedItem, bModify);
		}

		return nullptr;
	}

	void CRecipientsDlg::OnViewRecipientABEntry()
	{
		m_bViewRecipientABEntry = !m_bViewRecipientABEntry;
		OnRefreshView();
	}

	void CRecipientsDlg::OnDeleteSelectedItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMessage || !m_lpContentsTableListCtrl) return;

		const auto iNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

		if (iNumSelected && iNumSelected < MAXNewADRLIST)
		{
			auto lpAdrList = mapi::allocate<LPADRLIST>(CbNewADRLIST(iNumSelected));
			if (lpAdrList)
			{
				lpAdrList->cEntries = iNumSelected;

				for (auto iSelection = UINT{}; iSelection < iNumSelected; iSelection++)
				{
					const auto lpProp = mapi::allocate<LPSPropValue>(sizeof(SPropValue));
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
							output::DBGDeleteSelectedItem,
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
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMessage || !m_lpContentsTableListCtrl || !m_lpPropDisplay) return;

		if (1 != m_lpContentsTableListCtrl->GetSelectedCount()) return;

		if (!m_lpPropDisplay->IsModifiedPropVals()) return;

		ULONG cProps = 0;
		LPSPropValue lpProps = nullptr;
		EC_H_S(m_lpPropDisplay->GetDisplayedProps(&cProps, &lpProps));
		if (lpProps)
		{
			ADRLIST adrList = {};
			adrList.cEntries = 1;
			adrList.aEntries[0].ulReserved1 = 0;
			adrList.aEntries[0].cValues = cProps;

			ULONG ulSizeProps = NULL;
			auto hRes = EC_MAPI(ScCountProps(adrList.aEntries[0].cValues, lpProps, &ulSizeProps));

			if (SUCCEEDED(hRes))
			{
				adrList.aEntries[0].rgPropVals = mapi::allocate<LPSPropValue>(ulSizeProps);
				if (adrList.aEntries[0].rgPropVals)
				{
					hRes = EC_MAPI(ScCopyProps(
						adrList.aEntries[0].cValues, lpProps, adrList.aEntries[0].rgPropVals, &ulSizeProps));
				}
			}

			if (SUCCEEDED(hRes))
			{
				output::DebugPrintEx(
					output::DBGGeneric, CLASS, L"OnModifyRecipients", L"Committing changes for current selection\n");
				EC_MAPI_S(m_lpMessage->ModifyRecipients(MODRECIP_MODIFY, &adrList));
			}

			MAPIFreeBuffer(adrList.aEntries[0].rgPropVals);

			OnRefreshView();
		}
	}

	void CRecipientsDlg::OnRecipOptions()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMessage || !m_lpContentsTableListCtrl || !m_lpPropDisplay || !m_lpMapiObjects) return;

		if (1 != m_lpContentsTableListCtrl->GetSelectedCount()) return;

		ULONG cProps = 0;
		LPSPropValue lpProps = nullptr;
		auto hRes = EC_H(m_lpPropDisplay->GetDisplayedProps(&cProps, &lpProps));

		if (lpProps)
		{
			auto lpAB = m_lpMapiObjects->GetAddrBook(true); // do not release
			if (lpAB)
			{
				ADRENTRY adrEntry = {};
				adrEntry.ulReserved1 = 0;
				adrEntry.cValues = cProps;
				adrEntry.rgPropVals = lpProps;
				output::DebugPrintEx(output::DBGGeneric, CLASS, L"OnRecipOptions", L"Calling RecipOptions\n");

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
					ADRLIST adrList = {};
					adrList.cEntries = 1;
					adrList.aEntries[0].ulReserved1 = 0;
					adrList.aEntries[0].cValues = adrEntry.cValues;
					adrList.aEntries[0].rgPropVals = adrEntry.rgPropVals;

					const auto szAdrList = property::AdrListToString(adrList);

					output::DebugPrintEx(
						output::DBGGeneric,
						CLASS,
						L"OnRecipOptions",
						L"RecipOptions returned the following ADRLIST:\n");
					// Note - debug output may be truncated due to limitations of OutputDebugString,
					// but output to file is complete
					output::Output(output::DBGGeneric, nullptr, false, szAdrList);

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
} // namespace dialog
