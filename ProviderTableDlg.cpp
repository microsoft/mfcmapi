// ProviderTableDlg.cpp : implementation file
// Displays the list of providers in a message service in a profile

#include "stdafx.h"
#include "ProviderTableDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "MAPIProfileFunctions.h"
#include "Editor.h"
#include "MapiFunctions.h"

static TCHAR* CLASS = _T("CProviderTableDlg");

/////////////////////////////////////////////////////////////////////////////
// CProviderTableDlg dialog


CProviderTableDlg::CProviderTableDlg(
			   CParentWnd* pParentWnd,
			   CMapiObjects* lpMapiObjects,
			   LPMAPITABLE lpMAPITable,
			   LPPROVIDERADMIN lpProviderAdmin
			   ):
CContentsTableDlg(
						  pParentWnd,
						  lpMapiObjects,
						  IDS_PROVIDERS,
						  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
						  lpMAPITable,
						  (LPSPropTagArray) &sptPROVIDERCols,
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
	m_lpContentsTable = NULL;
	if (m_lpProviderAdmin) m_lpProviderAdmin->Release();
}

BEGIN_MESSAGE_MAP(CProviderTableDlg, CContentsTableDlg)
	ON_COMMAND(ID_OPENPROFILESECTION,OnOpenProfileSection)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProviderTableDlg message handlers

HRESULT CProviderTableDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum /*bModify*/, LPMAPIPROP* lppMAPIProp)
{
	HRESULT			hRes = S_OK;
	LPSBinary		lpProviderUID = NULL;
	SortListData*	lpListData = NULL;

	DebugPrintEx(DBGOpenItemProp,CLASS,_T("OpenItemProp"),_T("iSelectedItem = 0x%X\n"),iSelectedItem);

	if (!lppMAPIProp || !m_lpContentsTableListCtrl || !m_lpProviderAdmin) return MAPI_E_INVALID_PARAMETER;

	lpListData = (SortListData*) m_lpContentsTableListCtrl->GetItemData(iSelectedItem);
	if (lpListData)
	{
		lpProviderUID = lpListData->data.Contents.lpProviderUID;
		if (lpProviderUID)
		{
			EC_H(OpenProfileSection(
				m_lpProviderAdmin,
				lpProviderUID,
				(LPPROFSECT*) lppMAPIProp));
		}
	}
	return hRes;
} // CProviderTableDlg::OpenItemProp

void CProviderTableDlg::OnOpenProfileSection()
{
	HRESULT			hRes = S_OK;

	if (!m_lpProviderAdmin) return;

	CEditor MyUID(
		this,
		IDS_OPENPROFSECT,
		IDS_OPENPROFSECTPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyUID.InitSingleLineSz(0,IDS_MAPIUID,_T("0a0d020000000000c000000000000046"),false); // STRING_OK

	WC_H(MyUID.DisplayDialog());
	if (S_OK != hRes) return;

	SBinary MapiUID = {0};
	CString szTmp;
	szTmp = MyUID.GetString(0);
	ULONG ulStrLen = szTmp.GetLength();

	if (32 != ulStrLen) return;

	MapiUID.cb = 16;

	EC_H(MAPIAllocateBuffer(
		MapiUID.cb,
		(LPVOID*)&MapiUID.lpb));
	MyBinFromHex(
		(LPCTSTR) szTmp,
		MapiUID.lpb,
		MapiUID.cb);

	LPPROFSECT lpProfSect = NULL;
	EC_H(OpenProfileSection(
		m_lpProviderAdmin,
		&MapiUID,
		&lpProfSect));
	if (lpProfSect)
	{
		LPMAPIPROP lpTemp = NULL;
		EC_H(lpProfSect->QueryInterface(IID_IMAPIProp,(LPVOID*) &lpTemp));
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

	return;
} // CProviderTableDlg::OnOpenProfileSection

void CProviderTableDlg::HandleAddInMenuSingle(
									   LPADDINMENUPARAMS lpParams,
									   LPMAPIPROP lpMAPIProp,
									   LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpProfSect = (LPPROFSECT) lpMAPIProp;
	}

	InvokeAddInMenu(lpParams);
} // CProviderTableDlg::HandleAddInMenuSingle