// ProfileListDlg.cpp : implementation file
// Displays the list of profiles

#include "stdafx.h"
#include "ProfileListDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "MAPIProfileFunctions.h"
#include "FileDialogEx.h"
#include "Editor.h"
#include "MsgServiceTableDlg.h"
#include "ImportProcs.h"
#include "MAPIFunctions.h"

static TCHAR* CLASS = _T("CProfileListDlg");

/////////////////////////////////////////////////////////////////////////////
// CProfileListDlg dialog


CProfileListDlg::CProfileListDlg(
								 CParentWnd* pParentWnd,
								 CMapiObjects* lpMapiObjects,
								 LPMAPITABLE lpMAPITable
								 ):
CContentsTableDlg(
				  pParentWnd,
				  lpMapiObjects,
				  IDS_PROFILES,
				  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
				  lpMAPITable,
				  (LPSPropTagArray) &sptPROFLISTCols,
				  NUMPROFLISTCOLUMNS,
				  PROFLISTColumns,
				  IDR_MENU_PROFILE_POPUP,
				  MENU_CONTEXT_PROFILE_LIST)
{
	TRACE_CONSTRUCTOR(CLASS);

	CreateDialogAndMenu(IDR_MENU_PROFILE);

	// Wipe the reference to the contents table to get around profile table refresh problems
	if (m_lpContentsTable) m_lpContentsTable->Release();
	m_lpContentsTable = NULL;
}

CProfileListDlg::~CProfileListDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

BEGIN_MESSAGE_MAP(CProfileListDlg, CContentsTableDlg)
ON_COMMAND(ID_GETMAPISVCINF,OnGetMAPISVC)
ON_COMMAND(ID_ADDSERVICESTOMAPISVCINF,OnAddServicesToMAPISVC)
ON_COMMAND(ID_REMOVESERVICESFROMMAPISVCINF,OnRemoveServicesFromMAPISVC)
ON_COMMAND(ID_DELETESELECTEDITEM,OnDeleteSelectedItem)
ON_COMMAND(ID_LAUNCHPROFILEWIZARD,OnLaunchProfileWizard)
ON_COMMAND(ID_ADDEXCHANGETOPROFILE,OnAddExchangeToProfile)
ON_COMMAND(ID_ADDPSTTOPROFILE,OnAddPSTToProfile)
ON_COMMAND(ID_ADDUNICODEPSTTOPROFILE,OnAddUnicodePSTToProfile)
ON_COMMAND(ID_ADDSERVICETOPROFILE,OnAddServiceToProfile)
ON_COMMAND(ID_GETPROFILESERVERVERSION,OnGetProfileServiceVersion)
ON_COMMAND(ID_CREATEPROFILE,OnCreateProfile)
END_MESSAGE_MAP()

void CProfileListDlg::OnInitMenu(CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ADDEXCHANGETOPROFILE,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ADDPSTTOPROFILE,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ADDUNICODEPSTTOPROFILE,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ADDSERVICETOPROFILE,DIMMSOK(iNumSel));
		}
		pMenu->EnableMenuItem(ID_LAUNCHPROFILEWIZARD,DIM(pfnLaunchWizard));
	}
	CContentsTableDlg::OnInitMenu(pMenu);
}

/////////////////////////////////////////////////////////////////////////////
// CProfileListDlg message handlers

// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void CProfileListDlg::OnRefreshView()
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpProfTable = NULL;

	if (!m_lpContentsTableListCtrl || !m_lpMapiObjects) return;

	if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
	DebugPrintEx(DBGGeneric,CLASS,_T("OnRefreshView"),_T("\n"));

	// Wipe out current references to the profile table so the refresh will work
	// If we don't do this, we get the old table back again.
	EC_H(m_lpContentsTableListCtrl->SetContentsTable(
		NULL,
		dfNormal,
		NULL));

	LPPROFADMIN lpProfAdmin = m_lpMapiObjects->GetProfAdmin(); // do not release

	if (!lpProfAdmin) return;

	EC_H(lpProfAdmin->GetProfileTable(
		0, // fMapiUnicode is not supported
		&lpProfTable));

	if (lpProfTable)
	{
		EC_H(m_lpContentsTableListCtrl->SetContentsTable(
			lpProfTable,
			dfNormal,
			NULL));

		lpProfTable->Release();
	}
} // CProfileListDlg::OnRefreshView

void CProfileListDlg::OnDisplayItem()
{
	HRESULT			hRes = S_OK;
	CHAR*			szProfileName = NULL;
	int				iItem = -1;
	SortListData*	lpListData = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	do
	{
		hRes = S_OK;
		lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (!lpListData) break;

		szProfileName = lpListData->data.Contents.szProfileDisplayName;
		if (szProfileName)
		{
			new CMsgServiceTableDlg(
				m_lpParent,
				m_lpMapiObjects,
				szProfileName);
		}
	}
	while (iItem != -1);

	return;
} // CProfileListDlg::OnDisplayItem

void CProfileListDlg::OnLaunchProfileWizard()
{
	HRESULT hRes = S_OK;
	CEditor MyData(
		this,
		IDS_LAUNCHPROFWIZ,
		IDS_LAUNCHPROFWIZPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitSingleLine(0,IDS_FLAGS,NULL,false);
	MyData.SetHex(0,MAPI_PW_LAUNCHED_BY_CONFIG);
	MyData.InitSingleLineSz(1,IDS_SERVICE,_T("MSEMS"),false); // STRING_OK

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		TCHAR szProfName[80] = {0};
		LPSTR szServices[] = {MyData.GetStringA(1),NULL};
		LaunchProfileWizard(
			m_hWnd,
			MyData.GetHex(0),
			(LPCSTR FAR *) szServices,
			_countof(szProfName),
			szProfName);
		OnRefreshView(); // Update the view since we don't have notifications here.
	}
} // CProfileListDlg::OnLaunchProfileWizard

void CProfileListDlg::OnGetMAPISVC()
{
	DisplayMAPISVCPath(this);
} // CProfileListDlg::OnGetMAPISVC

void CProfileListDlg::OnAddServicesToMAPISVC()
{
	AddServicesToMapiSvcInf();
} // CProfileListDlg::OnAddServicesToMAPISVC

void CProfileListDlg::OnRemoveServicesFromMAPISVC()
{
	RemoveServicesFromMapiSvcInf();
} // CProfileListDlg::OnRemoveServicesFromMAPISVC

void CProfileListDlg::OnAddExchangeToProfile()
{
	HRESULT			hRes = S_OK;
	int				iItem = -1;
	SortListData*	lpListData = NULL;

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_NEWEXPROF,
		IDS_NEWEXPROFPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_SERVERNAME,NULL,false);
	MyData.InitSingleLine(1,IDS_MAILBOXNAME,NULL,false);

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		LPSTR szServer = MyData.GetStringA(0);
		LPSTR szMailbox = MyData.GetStringA(1);

		if (szServer && szMailbox)
		{
			do
			{
				hRes = S_OK;
				// Find the highlighted item
				lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
				if (!lpListData) break;

				DebugPrintEx(DBGGeneric,CLASS,
					_T("OnAddExchangeToProfile"), // STRING_OK
					_T("Adding Server \"%hs\" and Mailbox \"%hs\" to profile \"%hs\"\n"), // STRING_OK
					szServer,
					szMailbox,
					lpListData->data.Contents.szProfileDisplayName);

				EC_H(HrAddExchangeToProfile(
					(ULONG_PTR) m_hWnd,
					szServer,
					szMailbox,
					lpListData->data.Contents.szProfileDisplayName));
			}
			while (iItem != -1);
		}
	}
	return;
} // CProfileListDlg::OnAddExchangeToProfile

void CProfileListDlg::AddPSTToProfile(BOOL bUnicodePST)
{
	HRESULT			hRes = S_OK;
	int				iItem = -1;
	SortListData*	lpListData = NULL;
	INT_PTR			iDlgRet = IDOK;

	CString szFileSpec;
	szFileSpec.LoadString(IDS_PSTFILES);

	if (!m_lpContentsTableListCtrl) return;

	CFileDialogEx dlgFilePicker;
	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		TRUE,
		_T("pst"), // STRING_OK
		NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		szFileSpec,
		this));
	if (IDOK == iDlgRet)
	{
		do
		{
			hRes = S_OK;
			// Find the highlighted item
			lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
			if (!lpListData) break;

			CEditor MyFile(
				this,
				IDS_PSTPATH,
				IDS_PSTPATHPROMPT,
				3,
				CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
			MyFile.InitSingleLineSz(0,IDS_SERVICE,dlgFilePicker.GetFileName(),false);
			MyFile.InitCheck(1,IDS_PSTDOPW,false,false);
			MyFile.InitSingleLineSz(2,IDS_PSTPW,_T(""),false);

			WC_H(MyFile.DisplayDialog());

			if (S_OK == hRes)
			{
				LPTSTR szPath = MyFile.GetString(0);
				BOOL bPasswordSet = MyFile.GetCheck(1);
				LPSTR szPwd = MyFile.GetStringA(2);

				DebugPrintEx(DBGGeneric,CLASS,_T("AddPSTToProfile"),_T("Adding PST \"%s\" to profile \"%hs\", bUnicodePST = 0x%X\n, bPasswordSet = 0x%X, password = \"%hs\"\n"),
					szPath,
					lpListData->data.Contents.szProfileDisplayName,
					bUnicodePST,
					bPasswordSet,
					szPwd);

				CWaitCursor Wait; // Change the mouse to an hourglass while we work.
				EC_H(HrAddPSTToProfile((ULONG_PTR) m_hWnd,bUnicodePST,szPath,lpListData->data.Contents.szProfileDisplayName,bPasswordSet,szPwd));
			}
		}
		while (iItem != -1);
	}
	return;
} // CProfileListDlg::AddPSTToProfile

void CProfileListDlg::OnAddPSTToProfile()
{
	AddPSTToProfile(false);
} // CProfileListDlg::OnAddPSTToProfile

void CProfileListDlg::OnAddUnicodePSTToProfile()
{
	AddPSTToProfile(true);
} // CProfileListDlg::OnAddUnicodePSTToProfile

void CProfileListDlg::OnAddServiceToProfile()
{
	HRESULT			hRes = S_OK;
	int				iItem = -1;
	SortListData*	lpListData = NULL;

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_NEWSERVICE,
		IDS_NEWSERVICEPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitSingleLine(0,IDS_SERVICE,NULL,false);
	MyData.InitCheck(1,IDS_DISPLAYSERVICEUI,true,false);

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		LPSTR szService = MyData.GetStringA(0);
		do
		{
			hRes = S_OK;
			// Find the highlighted item
			lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
			if (!lpListData) break;

			DebugPrintEx(DBGGeneric,CLASS,_T("OnAddServiceToProfile"),_T("Adding service \"%hs\" to profile \"%hs\"\n"),szService,lpListData->data.Contents.szProfileDisplayName);

			EC_H(HrAddServiceToProfile(szService,(ULONG_PTR) m_hWnd,MyData.GetCheck(1)?SERVICE_UI_ALWAYS:0,0,0,lpListData->data.Contents.szProfileDisplayName));
		}
		while (iItem != -1);
	}

	return;
} // CProfileListDlg::OnAddServiceToProfile

void CProfileListDlg::OnCreateProfile()
{
	HRESULT			hRes = S_OK;

	CEditor MyData(
		this,
		IDS_NEWPROF,
		IDS_NEWPROFPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitSingleLine(0,IDS_PROFILE,NULL,false);

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		CWaitCursor		Wait; // Change the mouse to an hourglass while we work.
		LPSTR szProfile = MyData.GetStringA(0);

		DebugPrintEx(DBGGeneric,CLASS,
			_T("OnCreateProfile"), // STRING_OK
			_T("Creating profile \"%hs\"\n"), // STRING_OK
			szProfile);

		EC_H(HrCreateProfile(szProfile));

		// Since we may have created a profile, update even if we failed.
		OnRefreshView(); // Update the view since we don't have notifications here.
	}
	return;
} // CProfileListDlg::OnCreateProfile

void CProfileListDlg::OnDeleteSelectedItem()
{
	HRESULT		hRes = S_OK;
	int			iItem = -1;
	SortListData*	lpListData = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	do
	{
		hRes = S_OK;
		// Find the highlighted item AttachNum
		lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (!lpListData) break;

		DebugPrintEx(DBGDeleteSelectedItem,CLASS,_T("OnDeleteSelectedItem"),_T("Deleting profile \"%hs\"\n"),lpListData->data.Contents.szProfileDisplayName);

		EC_H(HrRemoveProfile(lpListData->data.Contents.szProfileDisplayName));
	}
	while (iItem != -1);

	OnRefreshView(); // Update the view since we don't have notifications here.
} // CProfileListDlg::OnDeleteSelectedItem

void CProfileListDlg::OnGetProfileServiceVersion()
{
	HRESULT		hRes = S_OK;
	int			iItem = -1;
	SortListData*	lpListData = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	do
	{
		hRes = S_OK;
		// Find the highlighted item AttachNum
		lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (!lpListData) break;

		DebugPrintEx(DBGDeleteSelectedItem,CLASS,_T("OnGetProfileServiceVersion"),_T("Getting profile service version for \"%hs\"\n"),lpListData->data.Contents.szProfileDisplayName);

		ULONG ulServerVersion = 0;
		EXCHANGE_STORE_VERSION_NUM storeVersion = {0};
		BOOL bFoundServerVersion = false;
		BOOL bFoundServerFullVersion = false;

		WC_H(GetProfileServiceVersion(lpListData->data.Contents.szProfileDisplayName,
			&ulServerVersion,
			&storeVersion,
			&bFoundServerVersion,
			&bFoundServerFullVersion));
		// Even in failure case, we're still gonna show the dialog
		hRes = S_OK;

		CEditor MyData(
			this,
			IDS_PROFILESERVERVERSIONTITLE,
			IDS_PROFILESERVERVERSIONPROMPT,
			5,
			CEDITOR_BUTTON_OK);

		MyData.InitSingleLine(0,IDS_PROFILESERVERVERSION,NULL,true);
		MyData.InitSingleLine(1,IDS_PROFILESERVERVERSIONMAJOR,NULL,true);
		MyData.InitSingleLine(2,IDS_PROFILESERVERVERSIONMINOR,NULL,true);
		MyData.InitSingleLine(3,IDS_PROFILESERVERVERSIONBUILD,NULL,true);
		MyData.InitSingleLine(4,IDS_PROFILESERVERVERSIONMINORBUILD,NULL,true);

		if (bFoundServerVersion)
		{
			MyData.SetStringf(0,_T("%d = 0x%X"),ulServerVersion,ulServerVersion); // STRING_OK
			DebugPrint(DBGGeneric,_T("PR_PROFILE_SERVER_VERSION == 0x%08X\n"),ulServerVersion);
		}
		else
		{
			MyData.LoadString(0,IDS_NOTFOUND);
		}

		if (bFoundServerFullVersion)
		{
			DebugPrint(DBGGeneric,_T("PR_PROFILE_SERVER_FULL_VERSION = \n"));
			MyData.SetStringf(1,_T("%d = 0x%X"),storeVersion.wMajorVersion,storeVersion.wMajorVersion); // STRING_OK
			MyData.SetStringf(2,_T("%d = 0x%X"),storeVersion.wMinorVersion,storeVersion.wMinorVersion); // STRING_OK
			MyData.SetStringf(3,_T("%d = 0x%X"),storeVersion.wBuild,storeVersion.wBuild); // STRING_OK
			MyData.SetStringf(4,_T("%d = 0x%X"),storeVersion.wMinorBuild,storeVersion.wMinorBuild); // STRING_OK
			DebugPrint(DBGGeneric,
				_T("\twMajorVersion 0x%04X\n") // STRING_OK
				_T("\twMinorVersion 0x%04X\n") // STRING_OK
				_T("\twBuild 0x%04X\n") // STRING_OK
				_T("\twMinorBuild 0x%04X\n"), // STRING_OK
				storeVersion.wMajorVersion,
				storeVersion.wMinorVersion,
				storeVersion.wBuild,
				storeVersion.wMinorBuild);
		}
		else
		{
			MyData.LoadString(1,IDS_NOTFOUND);
			MyData.LoadString(2,IDS_NOTFOUND);
			MyData.LoadString(3,IDS_NOTFOUND);
			MyData.LoadString(4,IDS_NOTFOUND);
		}
		WC_H(MyData.DisplayDialog());
	}
	while (iItem != -1);
} // CProfileListDlg::OnGetProfileServiceVersion