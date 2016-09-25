// ProfileListDlg.cpp : Displays the list of profiles

#include "stdafx.h"
#include "ProfileListDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "ColumnTags.h"
#include "MAPIProfileFunctions.h"
#include "FileDialogEx.h"
#include "Editor.h"
#include "MsgServiceTableDlg.h"
#include "String.h"
#include "File.h"
#include "ExportProfile.h"
#include "SortList/ContentsData.h"
#include "GlobalCache.h"

static wstring CLASS = L"CProfileListDlg";

CProfileListDlg::CProfileListDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPMAPITABLE lpMAPITable
) :
	CContentsTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_PROFILES,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		lpMAPITable,
		LPSPropTagArray(&sptPROFLISTCols),
		NUMPROFLISTCOLUMNS,
		PROFLISTColumns,
		IDR_MENU_PROFILE_POPUP,
		MENU_CONTEXT_PROFILE_LIST)
{
	TRACE_CONSTRUCTOR(CLASS);

	CreateDialogAndMenu(IDR_MENU_PROFILE);

	// Wipe the reference to the contents table to get around profile table refresh problems
	if (m_lpContentsTable) m_lpContentsTable->Release();
	m_lpContentsTable = nullptr;
}

CProfileListDlg::~CProfileListDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

BEGIN_MESSAGE_MAP(CProfileListDlg, CContentsTableDlg)
	ON_COMMAND(ID_GETMAPISVCINF, OnGetMAPISVC)
	ON_COMMAND(ID_ADDSERVICESTOMAPISVCINF, OnAddServicesToMAPISVC)
	ON_COMMAND(ID_REMOVESERVICESFROMMAPISVCINF, OnRemoveServicesFromMAPISVC)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_LAUNCHPROFILEWIZARD, OnLaunchProfileWizard)
	ON_COMMAND(ID_ADDEXCHANGETOPROFILE, OnAddExchangeToProfile)
	ON_COMMAND(ID_ADDPSTTOPROFILE, OnAddPSTToProfile)
	ON_COMMAND(ID_ADDUNICODEPSTTOPROFILE, OnAddUnicodePSTToProfile)
	ON_COMMAND(ID_ADDSERVICETOPROFILE, OnAddServiceToProfile)
	ON_COMMAND(ID_GETPROFILESERVERVERSION, OnGetProfileServiceVersion)
	ON_COMMAND(ID_CREATEPROFILE, OnCreateProfile)
	ON_COMMAND(ID_SETDEFAULTPROFILE, OnSetDefaultProfile)
	ON_COMMAND(ID_OPENPROFILEBYNAME, OnOpenProfileByName)
	ON_COMMAND(ID_EXPORTPROFILE, OnExportProfile)
END_MESSAGE_MAP()

void CProfileListDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ADDEXCHANGETOPROFILE, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ADDPSTTOPROFILE, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ADDUNICODEPSTTOPROFILE, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ADDSERVICETOPROFILE, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_SETDEFAULTPROFILE, DIMMSNOK(iNumSel));
			pMenu->EnableMenuItem(ID_EXPORTPROFILE, DIMMSNOK(iNumSel));

			pMenu->EnableMenuItem(ID_COPY, DIMMSNOK(iNumSel));
			auto ulStatus = CGlobalCache::getInstance().GetBufferStatus();
			pMenu->EnableMenuItem(ID_PASTE, DIM(ulStatus & BUFFER_PROFILE));
		}
	}

	CContentsTableDlg::OnInitMenu(pMenu);
}

// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void CProfileListDlg::OnRefreshView()
{
	auto hRes = S_OK;
	LPMAPITABLE lpProfTable = nullptr;

	if (!m_lpContentsTableListCtrl) return;

	if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
	DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L"\n");

	// Wipe out current references to the profile table so the refresh will work
	// If we don't do this, we get the old table back again.
	EC_H(m_lpContentsTableListCtrl->SetContentsTable(
		NULL,
		dfNormal,
		NULL));

	LPPROFADMIN lpProfAdmin = nullptr;
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
}

void CProfileListDlg::OnDisplayItem()
{
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	do
	{
		auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (!lpListData || !lpListData->Contents()) break;

		if (!lpListData->Contents()->m_szProfileDisplayName.empty())
		{
			new CMsgServiceTableDlg(
				m_lpParent,
				m_lpMapiObjects,
				lpListData->Contents()->m_szProfileDisplayName);
		}
	} while (iItem != -1);
}

void CProfileListDlg::OnLaunchProfileWizard()
{
	auto hRes = S_OK;
	CEditor MyData(
		this,
		IDS_LAUNCHPROFWIZ,
		IDS_LAUNCHPROFWIZPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGS, false));
	MyData.SetHex(0, MAPI_PW_LAUNCHED_BY_CONFIG);
	MyData.InitPane(1, CreateSingleLinePane(IDS_SERVICE, wstring(L"MSEMS"), false)); // STRING_OK

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		auto szProfName = LaunchProfileWizard(
			m_hWnd,
			MyData.GetHex(0),
			wstringTostring(MyData.GetStringW(1)));
		OnRefreshView(); // Update the view since we don't have notifications here.
	}
}

void CProfileListDlg::OnGetMAPISVC()
{
	DisplayMAPISVCPath(this);
}

void CProfileListDlg::OnAddServicesToMAPISVC()
{
	AddServicesToMapiSvcInf();
}

void CProfileListDlg::OnRemoveServicesFromMAPISVC()
{
	RemoveServicesFromMapiSvcInf();
}

void CProfileListDlg::OnAddExchangeToProfile()
{
	auto hRes = S_OK;
	auto iItem = -1;

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_NEWEXPROF,
		IDS_NEWEXPROFPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, false));
	MyData.InitPane(1, CreateSingleLinePane(IDS_MAILBOXNAME, false));

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto szServer = wstringTostring(MyData.GetStringW(0));
		auto szMailbox = wstringTostring(MyData.GetStringW(1));

		if (!szServer.empty() && !szMailbox.empty())
		{
			do
			{
				hRes = S_OK;
				// Find the highlighted item
				auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
				if (!lpListData || !lpListData->Contents()) break;

				DebugPrintEx(DBGGeneric, CLASS,
					L"OnAddExchangeToProfile", // STRING_OK
					L"Adding Server \"%hs\" and Mailbox \"%hs\" to profile \"%hs\"\n", // STRING_OK
					szServer.c_str(),
					szMailbox.c_str(),
					lpListData->Contents()->m_szProfileDisplayName.c_str());

				EC_H(HrAddExchangeToProfile(
					reinterpret_cast<ULONG_PTR>(m_hWnd),
					szServer,
					szMailbox,
					lpListData->Contents()->m_szProfileDisplayName));
			} while (iItem != -1);
		}
	}
}

void CProfileListDlg::AddPSTToProfile(bool bUnicodePST)
{
	auto hRes = S_OK;
	auto iItem = -1;
	INT_PTR iDlgRet = IDOK;

	auto szFileSpec = loadstring(IDS_PSTFILES);

	if (!m_lpContentsTableListCtrl) return;

	CFileDialogExW dlgFilePicker;
	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		true,
		L"pst", // STRING_OK
		emptystring,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		szFileSpec,
		this));
	if (IDOK == iDlgRet)
	{
		do
		{
			hRes = S_OK;
			// Find the highlighted item
			auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
			if (!lpListData || !lpListData->Contents()) break;

			CEditor MyFile(
				this,
				IDS_PSTPATH,
				IDS_PSTPATHPROMPT,
				3,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyFile.InitPane(0, CreateSingleLinePane(IDS_SERVICE, wstring(dlgFilePicker.GetFileName()), false));
			MyFile.InitPane(1, CreateCheckPane(IDS_PSTDOPW, false, false));
			MyFile.InitPane(2, CreateSingleLinePane(IDS_PSTPW, false));

			WC_H(MyFile.DisplayDialog());

			if (S_OK == hRes)
			{
				auto szPath = MyFile.GetStringW(0);
				auto bPasswordSet = MyFile.GetCheck(1);
				auto szPwd = wstringTostring(MyFile.GetStringW(2));

				DebugPrintEx(DBGGeneric, CLASS, L"AddPSTToProfile", L"Adding PST \"%ws\" to profile \"%hs\", bUnicodePST = 0x%X\n, bPasswordSet = 0x%X, password = \"%hs\"\n",
					szPath.c_str(),
					lpListData->Contents()->m_szProfileDisplayName.c_str(),
					bUnicodePST,
					bPasswordSet,
					szPwd.c_str());

				CWaitCursor Wait; // Change the mouse to an hourglass while we work.
				EC_H(HrAddPSTToProfile(reinterpret_cast<ULONG_PTR>(m_hWnd), bUnicodePST, szPath, lpListData->Contents()->m_szProfileDisplayName, bPasswordSet, szPwd));
			}
		} while (iItem != -1);
	}
}

void CProfileListDlg::OnAddPSTToProfile()
{
	AddPSTToProfile(false);
}

void CProfileListDlg::OnAddUnicodePSTToProfile()
{
	AddPSTToProfile(true);
}

void CProfileListDlg::OnAddServiceToProfile()
{
	auto hRes = S_OK;
	auto iItem = -1;

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_NEWSERVICE,
		IDS_NEWSERVICEPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_SERVICE, false));
	MyData.InitPane(1, CreateCheckPane(IDS_DISPLAYSERVICEUI, true, false));

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto szService = wstringTostring(MyData.GetStringW(0));
		do
		{
			hRes = S_OK;
			// Find the highlighted item
			auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
			if (!lpListData || !lpListData->Contents()) break;

			DebugPrintEx(DBGGeneric, CLASS, L"OnAddServiceToProfile", L"Adding service \"%hs\" to profile \"%hs\"\n", szService.c_str(), lpListData->Contents()->m_szProfileDisplayName.c_str());

			EC_H(HrAddServiceToProfile(szService, reinterpret_cast<ULONG_PTR>(m_hWnd), MyData.GetCheck(1) ? SERVICE_UI_ALWAYS : 0, 0, nullptr, lpListData->Contents()->m_szProfileDisplayName));
		} while (iItem != -1);
	}
}

void CProfileListDlg::OnCreateProfile()
{
	auto hRes = S_OK;

	CEditor MyData(
		this,
		IDS_NEWPROF,
		IDS_NEWPROFPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_PROFILE, false));

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto szProfile = wstringTostring(MyData.GetStringW(0));

		DebugPrintEx(DBGGeneric, CLASS,
			L"OnCreateProfile", // STRING_OK
			L"Creating profile \"%hs\"\n", // STRING_OK
			szProfile.c_str());

		EC_H(HrCreateProfile(szProfile));

		// Since we may have created a profile, update even if we failed.
		OnRefreshView(); // Update the view since we don't have notifications here.
	}
}

void CProfileListDlg::OnDeleteSelectedItem()
{
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	do
	{
		auto hRes = S_OK;
		// Find the highlighted item AttachNum
		auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (!lpListData || !lpListData->Contents()) break;

		DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L"Deleting profile \"%hs\"\n", lpListData->Contents()->m_szProfileDisplayName.c_str());

		EC_H(HrRemoveProfile(lpListData->Contents()->m_szProfileDisplayName));
	} while (iItem != -1);

	OnRefreshView(); // Update the view since we don't have notifications here.
}

void CProfileListDlg::OnGetProfileServiceVersion()
{
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	do
	{
		auto hRes = S_OK;
		// Find the highlighted item AttachNum
		auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (!lpListData || !lpListData->Contents()) break;

		DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnGetProfileServiceVersion", L"Getting profile service version for \"%hs\"\n", lpListData->Contents()->m_szProfileDisplayName.c_str());

		ULONG ulServerVersion = 0;
		EXCHANGE_STORE_VERSION_NUM storeVersion = { 0 };
		auto bFoundServerVersion = false;
		auto bFoundServerFullVersion = false;

		WC_H(GetProfileServiceVersion(
			lpListData->Contents()->m_szProfileDisplayName,
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

		MyData.InitPane(0, CreateSingleLinePane(IDS_PROFILESERVERVERSION, true));
		MyData.InitPane(1, CreateSingleLinePane(IDS_PROFILESERVERVERSIONMAJOR, true));
		MyData.InitPane(2, CreateSingleLinePane(IDS_PROFILESERVERVERSIONMINOR, true));
		MyData.InitPane(3, CreateSingleLinePane(IDS_PROFILESERVERVERSIONBUILD, true));
		MyData.InitPane(4, CreateSingleLinePane(IDS_PROFILESERVERVERSIONMINORBUILD, true));

		if (bFoundServerVersion)
		{
			MyData.SetStringf(0, L"%u = 0x%X", ulServerVersion, ulServerVersion); // STRING_OK
			DebugPrint(DBGGeneric, L"PR_PROFILE_SERVER_VERSION == 0x%08X\n", ulServerVersion);
		}
		else
		{
			MyData.LoadString(0, IDS_NOTFOUND);
		}

		if (bFoundServerFullVersion)
		{
			DebugPrint(DBGGeneric, L"PR_PROFILE_SERVER_FULL_VERSION = \n");
			MyData.SetStringf(1, L"%u = 0x%X", storeVersion.wMajorVersion, storeVersion.wMajorVersion); // STRING_OK
			MyData.SetStringf(2, L"%u = 0x%X", storeVersion.wMinorVersion, storeVersion.wMinorVersion); // STRING_OK
			MyData.SetStringf(3, L"%u = 0x%X", storeVersion.wBuild, storeVersion.wBuild); // STRING_OK
			MyData.SetStringf(4, L"%u = 0x%X", storeVersion.wMinorBuild, storeVersion.wMinorBuild); // STRING_OK
			DebugPrint(DBGGeneric,
				L"\twMajorVersion 0x%04X\n" // STRING_OK
				L"\twMinorVersion 0x%04X\n" // STRING_OK
				L"\twBuild 0x%04X\n" // STRING_OK
				L"\twMinorBuild 0x%04X\n", // STRING_OK
				storeVersion.wMajorVersion,
				storeVersion.wMinorVersion,
				storeVersion.wBuild,
				storeVersion.wMinorBuild);
		}
		else
		{
			MyData.LoadString(1, IDS_NOTFOUND);
			MyData.LoadString(2, IDS_NOTFOUND);
			MyData.LoadString(3, IDS_NOTFOUND);
			MyData.LoadString(4, IDS_NOTFOUND);
		}

		WC_H(MyData.DisplayDialog());
	} while (iItem != -1);
}

void CProfileListDlg::OnSetDefaultProfile()
{
	auto hRes = S_OK;
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	// Find the highlighted item AttachNum
	auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
	if (lpListData || !lpListData->Contents())
	{
		DebugPrintEx(DBGGeneric, CLASS, L"OnSetDefaultProfile", L"Setting profile \"%hs\" as default\n", lpListData->Contents()->m_szProfileDisplayName.c_str());

		EC_H(HrSetDefaultProfile(lpListData->Contents()->m_szProfileDisplayName));

		OnRefreshView(); // Update the view since we don't have notifications here.
	}
}

void CProfileListDlg::OnOpenProfileByName()
{
	auto hRes = S_OK;

	CEditor MyData(
		this,
		IDS_OPENPROFILE,
		NULL,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_OPENPROFILEPROMPT, false));

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		auto szProfileName = wstringTostring(MyData.GetStringW(0));
		if (!szProfileName.empty())
		{
			new CMsgServiceTableDlg(
				m_lpParent,
				m_lpMapiObjects,
				szProfileName);
		}
	}
}

void CProfileListDlg::HandleCopy()
{
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
	if (!m_lpContentsTableListCtrl) return;

	// Find the highlighted profile
	auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
	if (lpListData || !lpListData->Contents())
	{
		CGlobalCache::getInstance().SetProfileToCopy(lpListData->Contents()->m_szProfileDisplayName);
	}
}

_Check_return_ bool CProfileListDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");

	auto szOldProfile = CGlobalCache::getInstance().GetProfileToCopy();

	CEditor MyData(
		this,
		IDS_COPYPROFILE,
		NULL,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePane(IDS_COPYPROFILEPROMPT, stringTowstring(szOldProfile), false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		auto szNewProfile = wstringTostring(MyData.GetStringW(0));

		WC_MAPI(HrCopyProfile(szOldProfile, szNewProfile));

		OnRefreshView(); // Update the view since we don't have notifications here.
	}

	return true;
}

void CProfileListDlg::OnExportProfile()
{
	auto hRes = S_OK;
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"OnExportProfile", L"\n");
	if (!m_lpContentsTableListCtrl) return;

	// Find the highlighted profile
	auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
	if (lpListData && lpListData->Contents())
	{
		INT_PTR iDlgRet = IDOK;

		auto szProfileName = stringTowstring(lpListData->Contents()->m_szProfileDisplayName);
		auto szFileName = BuildFileNameAndPath(
			L".xml", // STRING_OK
			szProfileName,
			emptystring,
			nullptr);
		DebugPrint(DBGGeneric, L"BuildFileNameAndPath built file name \"%ws\"\n", szFileName.c_str());

		auto szFileSpec = loadstring(IDS_XMLFILES);
		CFileDialogExW dlgFilePicker;
		EC_D_DIALOG(dlgFilePicker.DisplayDialog(
			false,
			L"xml", // STRING_OK
			szFileName,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			szFileSpec.c_str(),
			this));

		if (iDlgRet == IDOK)
		{
			ExportProfile(lpListData->Contents()->m_szProfileDisplayName.c_str(), nullptr, false, dlgFilePicker.GetFileName());
		}
	}
}