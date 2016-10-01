// MsgServiceTableDlg.cpp : Displays the list of services in a profile

#include "stdafx.h"
#include "MsgServiceTableDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "ProviderTableDlg.h"
#include "MAPIProfileFunctions.h"
#include <Dialogs/Editors/Editor.h>
#include "SortList/ContentsData.h"

static wstring CLASS = L"CMsgServiceTableDlg";

CMsgServiceTableDlg::CMsgServiceTableDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ string szProfileName
) :
	CContentsTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_SERVICES,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		nullptr,
		LPSPropTagArray(&sptSERVICECols),
		NUMSERVICECOLUMNS,
		SERVICEColumns,
		IDR_MENU_MSGSERVICE_POPUP,
		MENU_CONTEXT_PROFILE_SERVICES)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_lpServiceAdmin = nullptr;

	CreateDialogAndMenu(IDR_MENU_MSGSERVICE);

	m_szProfileName = szProfileName;
	CMsgServiceTableDlg::OnRefreshView();
}

CMsgServiceTableDlg::~CMsgServiceTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	// little hack to keep our releases in the right order - crash in o2k3 otherwise
	if (m_lpContentsTable) m_lpContentsTable->Release();
	m_lpContentsTable = nullptr;
	if (m_lpServiceAdmin) m_lpServiceAdmin->Release();
}

BEGIN_MESSAGE_MAP(CMsgServiceTableDlg, CContentsTableDlg)
	ON_COMMAND(ID_CONFIGUREMSGSERVICE, OnConfigureMsgService)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_OPENPROFILESECTION, OnOpenProfileSection)
END_MESSAGE_MAP()

void CMsgServiceTableDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_CONFIGUREMSGSERVICE, DIMMSOK(iNumSel));
		}
	}

	CContentsTableDlg::OnInitMenu(pMenu);
}

// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void CMsgServiceTableDlg::OnRefreshView()
{
	DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L"\n");

	auto hRes = S_OK;

	// Make sure we've got something to work with
	if (m_szProfileName.empty() || !m_lpContentsTableListCtrl || !m_lpMapiObjects) return;

	// cancel any loading which may be occuring
	if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();

	// Clean up our table and admin in reverse order from which we obtained them
	// Failure to do this leads to crashes in Outlook's profile code
	EC_H(m_lpContentsTableListCtrl->SetContentsTable(
		NULL,
		dfNormal,
		NULL));

	if (m_lpServiceAdmin) m_lpServiceAdmin->Release();
	m_lpServiceAdmin = nullptr;

	LPPROFADMIN lpProfAdmin = nullptr;
	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));

	if (lpProfAdmin)
	{
		EC_MAPI(lpProfAdmin->AdminServices(
			reinterpret_cast<LPTSTR>(const_cast<LPSTR>(m_szProfileName.c_str())),
			reinterpret_cast<LPTSTR>(""),
			NULL,
			MAPI_DIALOG,
			&m_lpServiceAdmin));
		if (m_lpServiceAdmin)
		{
			LPMAPITABLE lpServiceTable = nullptr;

			EC_MAPI(m_lpServiceAdmin->GetMsgServiceTable(
				0, // fMapiUnicode is not supported
				&lpServiceTable));

			if (lpServiceTable)
			{
				EC_H(m_lpContentsTableListCtrl->SetContentsTable(
					lpServiceTable,
					dfNormal,
					NULL));

				lpServiceTable->Release();
			}
		}

		lpProfAdmin->Release();
	}
}

void CMsgServiceTableDlg::OnDisplayItem()
{
	auto hRes = S_OK;
	LPPROVIDERADMIN lpProviderAdmin = nullptr;
	LPMAPITABLE lpProviderTable = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl || !m_lpServiceAdmin) return;

	auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
	for (auto lpListData : items)
	{
		if (lpListData && lpListData->Contents())
		{
			auto lpServiceUID = lpListData->Contents()->m_lpServiceUID;
			if (lpServiceUID)
			{
				EC_MAPI(m_lpServiceAdmin->AdminProviders(
					reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb),
					0, // fMapiUnicode is not supported
					&lpProviderAdmin));

				if (lpProviderAdmin)
				{
					EC_MAPI(lpProviderAdmin->GetProviderTable(
						0, // fMapiUnicode is not supported
						&lpProviderTable));

					if (lpProviderTable)
					{
						new CProviderTableDlg(
							m_lpParent,
							m_lpMapiObjects,
							lpProviderTable,
							lpProviderAdmin);
						lpProviderTable->Release();
						lpProviderTable = nullptr;
					}

					lpProviderAdmin->Release();
					lpProviderAdmin = nullptr;
				}
			}
		}
	}
}

void CMsgServiceTableDlg::OnConfigureMsgService()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl || !m_lpServiceAdmin) return;

	auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
	for (auto lpListData : items)
	{
		if (lpListData && lpListData->Contents())
		{
			auto lpServiceUID = lpListData->Contents()->m_lpServiceUID;
			if (lpServiceUID)
			{
				EC_H_CANCEL(m_lpServiceAdmin->ConfigureMsgService(
					reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb),
					reinterpret_cast<ULONG_PTR>(m_hWnd),
					SERVICE_UI_ALWAYS,
					0,
					nullptr));
			}
		}

		hRes = S_OK;
	}
}

_Check_return_ HRESULT CMsgServiceTableDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum /*bModify*/, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	auto hRes = S_OK;

	DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

	*lppMAPIProp = nullptr;

	if (!m_lpServiceAdmin || !m_lpContentsTableListCtrl || !lppMAPIProp) return MAPI_E_INVALID_PARAMETER;

	auto lpListData = reinterpret_cast<SortListData*>(m_lpContentsTableListCtrl->GetItemData(iSelectedItem));
	if (lpListData && lpListData->Contents())
	{
		auto lpServiceUID = lpListData->Contents()->m_lpServiceUID;
		if (lpServiceUID)
		{
			EC_H(OpenProfileSection(
				m_lpServiceAdmin,
				lpServiceUID,
				reinterpret_cast<LPPROFSECT*>(lppMAPIProp)));
		}
	}

	return hRes;
}

void CMsgServiceTableDlg::OnOpenProfileSection()
{
	auto hRes = S_OK;

	if (!m_lpServiceAdmin) return;

	CEditor MyUID(
		this,
		IDS_OPENPROFSECT,
		IDS_OPENPROFSECTPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyUID.InitPane(0, CreateDropDownGuidPane(IDS_MAPIUID, false));
	MyUID.InitPane(1, CreateCheckPane(IDS_MAPIUIDBYTESWAPPED, false, false));

	WC_H(MyUID.DisplayDialog());
	if (S_OK != hRes) return;

	GUID guid = { 0 };
	SBinary MapiUID = { sizeof(GUID), reinterpret_cast<LPBYTE>(&guid) };
	(void)MyUID.GetSelectedGUID(0, MyUID.GetCheck(1), &guid);

	LPPROFSECT lpProfSect = nullptr;
	EC_H(OpenProfileSection(
		m_lpServiceAdmin,
		&MapiUID,
		&lpProfSect));
	if (lpProfSect)
	{
		LPMAPIPROP lpTemp = nullptr;
		EC_MAPI(lpProfSect->QueryInterface(IID_IMAPIProp, reinterpret_cast<LPVOID*>(&lpTemp)));
		if (lpTemp)
		{
			EC_H(DisplayObject(
				lpTemp,
				MAPI_PROFSECT,
				otContents,
				this));
			lpTemp->Release();
		}

		lpProfSect->Release();
	}

	MAPIFreeBuffer(MapiUID.lpb);
}

void CMsgServiceTableDlg::OnDeleteSelectedItem()
{
	auto hRes = S_OK;
	if (!m_lpServiceAdmin || !m_lpContentsTableListCtrl) return;

	auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
	for (auto lpListData : items)
	{
		// Find the highlighted item AttachNum
		if (!lpListData || !lpListData->Contents()) break;

		DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L"Deleting service from \"%hs\"\n", lpListData->Contents()->m_szProfileDisplayName.c_str());

		auto lpServiceUID = lpListData->Contents()->m_lpServiceUID;
		if (lpServiceUID)
		{
			WC_MAPI(m_lpServiceAdmin->DeleteMsgService(
				reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb)));
		}

	}

	OnRefreshView(); // Update the view since we don't have notifications here.
}

void CMsgServiceTableDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_ LPMAPIPROP lpMAPIProp,
	_In_ LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpProfSect = static_cast<LPPROFSECT>(lpMAPIProp); // OpenItemProp returns LPPROFSECT
	}

	InvokeAddInMenu(lpParams);
}
