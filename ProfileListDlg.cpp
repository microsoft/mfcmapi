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
#include "File.h"

static TCHAR* CLASS = _T("CProfileListDlg");

/////////////////////////////////////////////////////////////////////////////
// CProfileListDlg dialog


CProfileListDlg::CProfileListDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPMAPITABLE lpMAPITable
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
} // CProfileListDlg::CProfileListDlg

CProfileListDlg::~CProfileListDlg()
{
	TRACE_DESTRUCTOR(CLASS);
} // CProfileListDlg::~CProfileListDlg

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
	ON_COMMAND(ID_SETDEFAULTPROFILE,OnSetDefaultProfile)
	ON_COMMAND(ID_OPENPROFILEBYNAME,OnOpenProfileByName)
	ON_COMMAND(ID_EXPORTPROFILE,OnExportProfile)
END_MESSAGE_MAP()

void CProfileListDlg::OnInitMenu(_In_ CMenu* pMenu)
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
			pMenu->EnableMenuItem(ID_SETDEFAULTPROFILE,DIMMSNOK(iNumSel));
			pMenu->EnableMenuItem(ID_EXPORTPROFILE,DIMMSNOK(iNumSel));

			pMenu->EnableMenuItem(ID_COPY,DIMMSNOK(iNumSel));
			ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();
			pMenu->EnableMenuItem(ID_PASTE,DIM(ulStatus & BUFFER_PROFILE));
		}
	}
	CContentsTableDlg::OnInitMenu(pMenu);
} // CProfileListDlg::OnInitMenu

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

	LPPROFADMIN lpProfAdmin = NULL;
	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return;

	EC_MAPI(lpProfAdmin->GetProfileTable(
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
	lpProfAdmin->Release();
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
	MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGS, NULL, false));
	MyData.SetHex(0,MAPI_PW_LAUNCHED_BY_CONFIG);
	MyData.InitPane(1, CreateSingleLinePane(IDS_SERVICE, _T("MSEMS"), false)); // STRING_OK

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		CHAR szProfName[80] = {0};
		LPSTR szServices[] = {MyData.GetStringA(1),NULL};
		LaunchProfileWizard(
			m_hWnd,
			MyData.GetHex(0),
			(LPCSTR*) szServices,
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

	MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, NULL, false));
	MyData.InitPane(1, CreateSingleLinePane(IDS_MAILBOXNAME, NULL, false));

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
} // CProfileListDlg::OnAddExchangeToProfile

void CProfileListDlg::AddPSTToProfile(bool bUnicodePST)
{
	HRESULT			hRes = S_OK;
	int				iItem = -1;
	SortListData*	lpListData = NULL;
	INT_PTR			iDlgRet = IDOK;

	CString szFileSpec;
	EC_B(szFileSpec.LoadString(IDS_PSTFILES));

	if (!m_lpContentsTableListCtrl) return;

	CFileDialogEx dlgFilePicker;
	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		true,
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
			MyFile.InitPane(0, CreateSingleLinePane(IDS_SERVICE, dlgFilePicker.GetFileName(), false));
			MyFile.InitPane(1, CreateCheckPane(IDS_PSTDOPW, false, false));
			MyFile.InitPane(2, CreateSingleLinePane(IDS_PSTPW, _T(""), false));

			WC_H(MyFile.DisplayDialog());

			if (S_OK == hRes)
			{
				LPTSTR szPath = MyFile.GetString(0);
				bool bPasswordSet = MyFile.GetCheck(1);
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
	MyData.InitPane(0, CreateSingleLinePane(IDS_SERVICE, NULL, false));
	MyData.InitPane(1, CreateCheckPane(IDS_DISPLAYSERVICEUI, true, false));

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
	MyData.InitPane(0, CreateSingleLinePane(IDS_PROFILE, NULL, false));

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
		bool bFoundServerVersion = false;
		bool bFoundServerFullVersion = false;

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

		MyData.InitPane(0, CreateSingleLinePane(IDS_PROFILESERVERVERSION, NULL, true));
		MyData.InitPane(1, CreateSingleLinePane(IDS_PROFILESERVERVERSIONMAJOR, NULL, true));
		MyData.InitPane(2, CreateSingleLinePane(IDS_PROFILESERVERVERSIONMINOR, NULL, true));
		MyData.InitPane(3, CreateSingleLinePane(IDS_PROFILESERVERVERSIONBUILD, NULL, true));
		MyData.InitPane(4, CreateSingleLinePane(IDS_PROFILESERVERVERSIONMINORBUILD, NULL, true));

		if (bFoundServerVersion)
		{
			MyData.SetStringf(0,_T("%u = 0x%X"),ulServerVersion,ulServerVersion); // STRING_OK
			DebugPrint(DBGGeneric,_T("PR_PROFILE_SERVER_VERSION == 0x%08X\n"),ulServerVersion);
		}
		else
		{
			MyData.LoadString(0,IDS_NOTFOUND);
		}

		if (bFoundServerFullVersion)
		{
			DebugPrint(DBGGeneric,_T("PR_PROFILE_SERVER_FULL_VERSION = \n"));
			MyData.SetStringf(1,_T("%u = 0x%X"),storeVersion.wMajorVersion,storeVersion.wMajorVersion); // STRING_OK
			MyData.SetStringf(2,_T("%u = 0x%X"),storeVersion.wMinorVersion,storeVersion.wMinorVersion); // STRING_OK
			MyData.SetStringf(3,_T("%u = 0x%X"),storeVersion.wBuild,storeVersion.wBuild); // STRING_OK
			MyData.SetStringf(4,_T("%u = 0x%X"),storeVersion.wMinorBuild,storeVersion.wMinorBuild); // STRING_OK
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

void CProfileListDlg::OnSetDefaultProfile()
{
	HRESULT hRes = S_OK;
	int iItem = -1;
	SortListData* lpListData = NULL;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	// Find the highlighted item AttachNum
	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
	if (lpListData)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("OnSetDefaultProfile"),_T("Setting profile \"%hs\" as default\n"),lpListData->data.Contents.szProfileDisplayName);

		EC_H(HrSetDefaultProfile(lpListData->data.Contents.szProfileDisplayName));

		OnRefreshView(); // Update the view since we don't have notifications here.
	}
} // CProfileListDlg::OnSetDefaultProfile

void CProfileListDlg::OnOpenProfileByName()
{
	HRESULT hRes = S_OK;
	CHAR* szProfileName = NULL;

	CEditor MyData(
		this,
		IDS_OPENPROFILE,
		NULL,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_OPENPROFILEPROMPT, NULL, false));

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		szProfileName = MyData.GetStringA(0);
		if (szProfileName)
		{
			new CMsgServiceTableDlg(
				m_lpParent,
				m_lpMapiObjects,
				szProfileName);
		}
	}
} // CProfileListDlg::OnOpenProfileByName

void CProfileListDlg::HandleCopy()
{
	int iItem = -1;
	SortListData* lpListData = NULL;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("HandleCopy"),_T("\n"));
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	// Find the highlighted profile
	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
	if (lpListData)
	{
		m_lpMapiObjects->SetProfileToCopy(lpListData->data.Contents.szProfileDisplayName);
	}
} // CProfileListDlg::HandleCopy

_Check_return_ bool CProfileListDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	HRESULT hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("HandlePaste"),_T("\n"));
	if (!m_lpMapiObjects) return false;

	LPSTR szOldProfile = m_lpMapiObjects->GetProfileToCopy();

	CEditor MyData(
		this,
		IDS_COPYPROFILE,
		NULL,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePaneA(IDS_COPYPROFILEPROMPT, szOldProfile, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPSTR szNewProfile = MyData.GetStringA(0);

		WC_MAPI(HrCopyProfile(szOldProfile, szNewProfile));

		OnRefreshView(); // Update the view since we don't have notifications here.
	}
	return true;
} // CProfileListDlg::HandlePaste

void ExportProfileSection(FILE* fProfile, LPPROFSECT lpSect)
{
	if (!fProfile || !lpSect) return;

	HRESULT hRes = S_OK;
	LPSPropValue lpAllProps = NULL;
	ULONG cValues = 0L;

	WC_H_GETPROPS(GetPropsNULL(lpSect,
		fMapiUnicode,
		&cValues,
		&lpAllProps));
	if (FAILED(hRes))
	{
		OutputToFilef(fProfile, _T("<properties error=\"0x%08X\" />\n"), hRes);
	}
	else if (lpAllProps)
	{
		OutputToFile(fProfile, _T("<properties listtype=\"section\">\n"));

		OutputPropertiesToFile(fProfile, cValues, lpAllProps, NULL, false);

		OutputToFile(fProfile, _T("</properties>\n"));

		MAPIFreeBuffer(lpAllProps);
	}
}

void ExportProfileProvider(FILE* fProfile, int iRow, LPPROVIDERADMIN lpProviderAdmin, LPSRow lpRow)
{
	if (!fProfile || !lpRow) return;

	Outputf(DBGNoDebug, fProfile, true, _T("<provider index = \"0x%08X\">\n"), iRow);

	OutputToFile(fProfile, _T("<properties listtype=\"row\">\n"));
	OutputSRowToFile(fProfile, lpRow, NULL);
	OutputToFile(fProfile, _T("</properties>\n"));

	HRESULT hRes = S_OK;
	LPSPropValue lpProviderUID = NULL;

	lpProviderUID = PpropFindProp(
		lpRow->lpProps,
		lpRow->cValues,
		PR_PROVIDER_UID);

	if (lpProviderUID)
	{
		LPPROFSECT lpSect = NULL;
		EC_H(OpenProfileSection(
			lpProviderAdmin,
			&lpProviderUID->Value.bin,
			&lpSect));
		if (lpSect)
		{
			ExportProfileSection(fProfile, lpSect);
			lpSect->Release();
		}
	}

	OutputToFile(fProfile, _T("</provider>\n"));
}

void ExportProfileService(FILE* fProfile, int iRow, LPSERVICEADMIN lpServiceAdmin, LPSRow lpRow)
{
	if (!fProfile || !lpRow) return;

	Outputf(DBGNoDebug, fProfile, true, _T("<service index = \"0x%08X\">\n"), iRow);

	OutputToFile(fProfile, _T("<properties listtype=\"row\">\n"));
	OutputSRowToFile(fProfile, lpRow, NULL);
	OutputToFile(fProfile, _T("</properties>\n"));

	HRESULT hRes = S_OK;
	LPSPropValue lpServiceUID = NULL;

	lpServiceUID = PpropFindProp(
		lpRow->lpProps,
		lpRow->cValues,
		PR_SERVICE_UID);

	if (lpServiceUID)
	{
		LPPROFSECT lpSect = NULL;
		EC_H(OpenProfileSection(
			lpServiceAdmin,
			&lpServiceUID->Value.bin,
			&lpSect));
		if (lpSect)
		{
			ExportProfileSection(fProfile, lpSect);
			lpSect->Release();
		}

		LPPROVIDERADMIN lpProviderAdmin = NULL;

		EC_MAPI(lpServiceAdmin->AdminProviders(
			(LPMAPIUID) lpServiceUID->Value.bin.lpb,
			0, // fMapiUnicode is not supported
			&lpProviderAdmin));

		if (lpProviderAdmin)
		{
			LPMAPITABLE lpProviderTable = NULL;
			EC_MAPI(lpProviderAdmin->GetProviderTable(
				0, // fMapiUnicode is not supported
				&lpProviderTable));

			if (lpProviderTable)
			{
				LPSRowSet lpRowSet = NULL;
				EC_MAPI(HrQueryAllRows(lpProviderTable, NULL, NULL, NULL, 0, &lpRowSet));
				if (lpRowSet && lpRowSet->cRows >= 1)
				{
					for (ULONG i = 0 ; i < lpRowSet->cRows ; i++)
					{
						ExportProfileProvider(fProfile, i, lpProviderAdmin, &lpRowSet->aRow[i]);
					}
				}

				FreeProws(lpRowSet);
				lpProviderTable->Release();
				lpProviderTable = NULL;
			}

			lpProviderAdmin->Release();
			lpProviderAdmin = NULL;
		}
	}

	OutputToFile(fProfile, _T("</service>\n"));
}

void ExportProfile(LPCSTR szProfile, LPWSTR szFileName)
{
	if (!szProfile || !szFileName) return;

	DebugPrint(DBGGeneric,_T("ExportProfile: Saving profile \"%hs\" to \"%ws\"\n"), szProfile, szFileName);

	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;
	FILE* fProfile = NULL;
	fProfile = MyOpenFile(szFileName, true);
	if (fProfile)
	{
		OutputToFile(fProfile, g_szXMLHeader);
		Outputf(DBGNoDebug, fProfile, true, _T("<profile profilename= \"%hs\">\n"), szProfile);

		EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));

		if (lpProfAdmin)
		{
			LPSERVICEADMIN lpServiceAdmin = NULL;
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
			EC_MAPI(lpProfAdmin->AdminServices(
				(TCHAR*)szProfile,
				(TCHAR*)"",
				NULL,
				MAPI_DIALOG,
				&lpServiceAdmin));
#pragma warning(pop)
			if (lpServiceAdmin)
			{
				LPMAPITABLE lpServiceTable = NULL;

				EC_MAPI(lpServiceAdmin->GetMsgServiceTable(
					0, // fMapiUnicode is not supported
					&lpServiceTable));

				if (lpServiceTable)
				{
					LPSRowSet lpRowSet = NULL;
					EC_MAPI(HrQueryAllRows(lpServiceTable, NULL, NULL, NULL, 0, &lpRowSet));
					if (lpRowSet && lpRowSet->cRows >= 1)
					{
						for (ULONG i = 0 ; i < lpRowSet->cRows ; i++)
						{
							ExportProfileService(fProfile, i, lpServiceAdmin, &lpRowSet->aRow[i]);
						}
					}

					FreeProws(lpRowSet);
					lpServiceTable->Release();
				}
				lpServiceAdmin->Release();
			}
			lpProfAdmin->Release();
		}

		OutputToFile(fProfile, _T("</profile>"));
		CloseFile(fProfile);
	}
}

void CProfileListDlg::OnExportProfile()
{
	HRESULT hRes = S_OK;
	int iItem = -1;
	SortListData* lpListData = NULL;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, _T("OnExportProfile"), _T("\n"));
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	// Find the highlighted profile
	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
	if (lpListData)
	{
		WCHAR szFileName[MAX_PATH] = {0};
		INT_PTR iDlgRet = IDOK;

		CStringW szFileSpec;
		EC_B(szFileSpec.LoadString(IDS_XMLFILES));

		LPWSTR szProfileName = NULL;
		EC_H(AnsiToUnicode(lpListData->data.Contents.szProfileDisplayName, &szProfileName));
		if (SUCCEEDED(hRes))
		{
			WC_H(BuildFileNameAndPath(
				szFileName,
				_countof(szFileName),
				L".xml", // STRING_OK
				4,
				szProfileName,
				NULL,
				NULL));
		}
		delete[] szProfileName;

		DebugPrint(DBGGeneric, _T("BuildFileNameAndPath built file name \"%ws\"\n"), szFileName);

		CFileDialogExW dlgFilePicker;

		EC_D_DIALOG(dlgFilePicker.DisplayDialog(
			false,
			L"xml", // STRING_OK
			szFileName,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			szFileSpec,
			this));

		if (iDlgRet == IDOK)
		{
			ExportProfile(lpListData->data.Contents.szProfileDisplayName, dlgFilePicker.GetFileName());
		}
	}
}