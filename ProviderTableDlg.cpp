// ProviderTableDlg.cpp : Displays the list of providers in a message service in a profile

#include "stdafx.h"
#include "ProviderTableDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "MAPIProfileFunctions.h"
#include "Editor.h"
#include "SortList/ContentsData.h"

static wstring CLASS = L"CProviderTableDlg";

CProviderTableDlg::CProviderTableDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPMAPITABLE lpMAPITable,
	_In_ LPPROVIDERADMIN lpProviderAdmin
) :
	CContentsTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_PROVIDERS,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		lpMAPITable,
		LPSPropTagArray(&sptPROVIDERCols),
		NUMPROVIDERCOLUMNS,
		PROVIDERColumns,
		NULL,
		MENU_CONTEXT_PROFILE_PROVIDERS)
{
	TRACE_CONSTRUCTOR(CLASS);

	CreateDialogAndMenu(IDR_MENU_PROVIDER);

	m_lpProviderAdmin = lpProviderAdmin;
	if (m_lpProviderAdmin) m_lpProviderAdmin->AddRef();
}

CProviderTableDlg::~CProviderTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	// little hack to keep our releases in the right order - crash in o2k3 otherwise
	if (m_lpContentsTable) m_lpContentsTable->Release();
	m_lpContentsTable = nullptr;
	if (m_lpProviderAdmin) m_lpProviderAdmin->Release();
}

BEGIN_MESSAGE_MAP(CProviderTableDlg, CContentsTableDlg)
	ON_COMMAND(ID_OPENPROFILESECTION, OnOpenProfileSection)
END_MESSAGE_MAP()

_Check_return_ HRESULT CProviderTableDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum /*bModify*/, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	auto hRes = S_OK;

	DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

	if (!lppMAPIProp || !m_lpContentsTableListCtrl || !m_lpProviderAdmin) return MAPI_E_INVALID_PARAMETER;

	*lppMAPIProp = nullptr;

	auto lpListData = reinterpret_cast<SortListData*>(m_lpContentsTableListCtrl->GetItemData(iSelectedItem));
	if (lpListData && lpListData->Contents())
	{
		auto lpProviderUID = lpListData->Contents()->m_lpProviderUID;
		if (lpProviderUID)
		{
			EC_H(OpenProfileSection(
				m_lpProviderAdmin,
				lpProviderUID,
				reinterpret_cast<LPPROFSECT*>(lppMAPIProp)));
		}
	}

	return hRes;
}

void CProviderTableDlg::OnOpenProfileSection()
{
	auto hRes = S_OK;

	if (!m_lpProviderAdmin) return;

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
		m_lpProviderAdmin,
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

void CProviderTableDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_ LPMAPIPROP lpMAPIProp,
	_In_ LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpProfSect = static_cast<LPPROFSECT>(lpMAPIProp);
	}

	InvokeAddInMenu(lpParams);
}