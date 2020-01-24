// Displays the list of profiles
#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/ProfileListDlg.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/columnTags.h>
#include <core/mapi/mapiProfileFunctions.h>
#include <UI/profile.h>
#include <UI/FileDialogEx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/ContentsTable/MsgServiceTableDlg.h>
#include <core/mapi/cache/globalCache.h>
#include <core/mapi/exportProfile.h>
#include <core/sortlistdata/sortListData.h>
#include <core/sortlistdata/contentsData.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace dialog
{
	static std::wstring CLASS = L"CProfileListDlg";

	CProfileListDlg::CProfileListDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ LPMAPITABLE lpMAPITable)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_PROFILES,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  lpMAPITable,
			  &columns::sptPROFLISTCols.tags,
			  columns::PROFLISTColumns,
			  IDR_MENU_PROFILE_POPUP,
			  MENU_CONTEXT_PROFILE_LIST)
	{
		TRACE_CONSTRUCTOR(CLASS);

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_PROFILE);

		// Wipe the reference to the contents table to get around profile table refresh problems
		if (m_lpContentsTable) m_lpContentsTable->Release();
		m_lpContentsTable = nullptr;
	}

	CProfileListDlg::~CProfileListDlg() { TRACE_DESTRUCTOR(CLASS); }

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
				const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
				pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_ADDEXCHANGETOPROFILE, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_ADDPSTTOPROFILE, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_ADDUNICODEPSTTOPROFILE, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_ADDSERVICETOPROFILE, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_SETDEFAULTPROFILE, DIMMSNOK(iNumSel));
				pMenu->EnableMenuItem(ID_EXPORTPROFILE, DIMMSNOK(iNumSel));

				pMenu->EnableMenuItem(ID_COPY, DIMMSNOK(iNumSel));
				const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();
				pMenu->EnableMenuItem(ID_PASTE, DIM(ulStatus & BUFFER_PROFILE));
			}
		}

		CContentsTableDlg::OnInitMenu(pMenu);
	}

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	void CProfileListDlg::OnRefreshView()
	{
		LPMAPITABLE lpProfTable = nullptr;

		if (!m_lpContentsTableListCtrl) return;

		if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
		output::DebugPrintEx(output::DBGGeneric, CLASS, L"OnRefreshView", L"\n");

		// Wipe out current references to the profile table so the refresh will work
		// If we don't do this, we get the old table back again.
		m_lpContentsTableListCtrl->SetContentsTable(nullptr, dfNormal, NULL);

		LPPROFADMIN lpProfAdmin = nullptr;
		EC_MAPI_S(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return;

		EC_MAPI_S(lpProfAdmin->GetProfileTable(
			0, // fMapiUnicode is not supported
			&lpProfTable));

		if (lpProfTable)
		{
			m_lpContentsTableListCtrl->SetContentsTable(lpProfTable, dfNormal, NULL);

			lpProfTable->Release();
		}

		lpProfAdmin->Release();
	}

	void CProfileListDlg::OnDisplayItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;
		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (!lpListData) break;
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (!contents) break;

			if (!contents->m_szProfileDisplayName.empty())
			{
				new CMsgServiceTableDlg(m_lpParent, m_lpMapiObjects, contents->m_szProfileDisplayName);
			}
		}
	}

	void CProfileListDlg::OnLaunchProfileWizard()
	{
		editor::CEditor MyData(
			this, IDS_LAUNCHPROFWIZ, IDS_LAUNCHPROFWIZPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLAGS, false));
		MyData.SetHex(0, MAPI_PW_LAUNCHED_BY_CONFIG);
		MyData.AddPane(
			viewpane::TextPane::CreateSingleLinePane(1, IDS_SERVICE, std::wstring(L"MSEMS"), false)); // STRING_OK

		if (!MyData.DisplayDialog()) return;

		auto szProfName = ui::profile::LaunchProfileWizard(m_hWnd, MyData.GetHex(0), MyData.GetStringW(1));
		OnRefreshView(); // Update the view since we don't have notifications here.
	}

	void CProfileListDlg::OnGetMAPISVC() { ui::profile::DisplayMAPISVCPath(this); }

	void CProfileListDlg::OnAddServicesToMAPISVC() { ui::profile::AddServicesToMapiSvcInf(); }

	void CProfileListDlg::OnRemoveServicesFromMAPISVC() { ui::profile::RemoveServicesFromMapiSvcInf(); }

	void CProfileListDlg::OnAddExchangeToProfile()
	{
		if (!m_lpContentsTableListCtrl) return;

		editor::CEditor MyData(this, IDS_NEWEXPROF, IDS_NEWEXPROFPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SERVERNAME, false));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_MAILBOXNAME, false));

		if (!MyData.DisplayDialog()) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto szServer = MyData.GetStringW(0);
		auto szMailbox = MyData.GetStringW(1);

		if (!szServer.empty() && !szMailbox.empty())
		{
			auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
			for (const auto& lpListData : items)
			{
				if (!lpListData) break;
				const auto contents = lpListData->cast<sortlistdata::contentsData>();
				if (!contents) break;

				output::DebugPrintEx(
					output::DBGGeneric,
					CLASS,
					L"OnAddExchangeToProfile", // STRING_OK
					L"Adding Server \"%ws\" and Mailbox \"%ws\" to profile \"%ws\"\n", // STRING_OK
					szServer.c_str(),
					szMailbox.c_str(),
					contents->m_szProfileDisplayName.c_str());

				EC_H_S(mapi::profile::HrAddExchangeToProfile(
					reinterpret_cast<ULONG_PTR>(m_hWnd), szServer, szMailbox, contents->m_szProfileDisplayName));
			}
		}
	}

	void CProfileListDlg::AddPSTToProfile(bool bUnicodePST)
	{
		if (!m_lpContentsTableListCtrl) return;

		auto file = file::CFileDialogExW::OpenFile(
			L"pst", // STRING_OK
			strings::emptystring,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			strings::loadstring(IDS_PSTFILES),
			this);
		if (!file.empty())
		{
			auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
			for (const auto& lpListData : items)
			{
				if (!lpListData) break;
				const auto contents = lpListData->cast<sortlistdata::contentsData>();
				if (!contents) break;

				editor::CEditor MyFile(this, IDS_PSTPATH, IDS_PSTPATHPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
				MyFile.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SERVICE, file, false));
				MyFile.AddPane(viewpane::CheckPane::Create(1, IDS_PSTDOPW, false, false));
				MyFile.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_PSTPW, false));

				if (MyFile.DisplayDialog())
				{
					auto szPath = MyFile.GetStringW(0);
					const auto bPasswordSet = MyFile.GetCheck(1);
					auto szPwd = MyFile.GetStringW(2);

					output::DebugPrintEx(
						output::DBGGeneric,
						CLASS,
						L"AddPSTToProfile",
						L"Adding PST \"%ws\" to profile \"%ws\", bUnicodePST = 0x%X\n, bPasswordSet = 0x%X, password = "
						L"\"%ws\"\n",
						szPath.c_str(),
						contents->m_szProfileDisplayName.c_str(),
						bUnicodePST,
						bPasswordSet,
						szPwd.c_str());

					CWaitCursor Wait; // Change the mouse to an hourglass while we work.
					EC_H_S(mapi::profile::HrAddPSTToProfile(
						reinterpret_cast<ULONG_PTR>(m_hWnd),
						bUnicodePST,
						szPath,
						contents->m_szProfileDisplayName,
						bPasswordSet,
						szPwd));
				}
			}
		}
	}

	void CProfileListDlg::OnAddPSTToProfile() { AddPSTToProfile(false); }

	void CProfileListDlg::OnAddUnicodePSTToProfile() { AddPSTToProfile(true); }

	void CProfileListDlg::OnAddServiceToProfile()
	{
		if (!m_lpContentsTableListCtrl) return;

		editor::CEditor MyData(this, IDS_NEWSERVICE, IDS_NEWSERVICEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SERVICE, false));
		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_DISPLAYSERVICEUI, true, false));

		if (!MyData.DisplayDialog()) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto szService = MyData.GetStringW(0);
		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (!lpListData) break;
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (!contents) break;

			output::DebugPrintEx(
				output::DBGGeneric,
				CLASS,
				L"OnAddServiceToProfile",
				L"Adding service \"%ws\" to profile \"%ws\"\n",
				szService.c_str(),
				contents->m_szProfileDisplayName.c_str());

			EC_H_S(mapi::profile::HrAddServiceToProfile(
				szService,
				reinterpret_cast<ULONG_PTR>(m_hWnd),
				MyData.GetCheck(1) ? SERVICE_UI_ALWAYS : 0,
				0,
				nullptr,
				contents->m_szProfileDisplayName));
		}
	}

	void CProfileListDlg::OnCreateProfile()
	{
		editor::CEditor MyData(this, IDS_NEWPROF, IDS_NEWPROFPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_PROFILE, false));

		if (!MyData.DisplayDialog()) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto szProfile = MyData.GetStringW(0);

		output::DebugPrintEx(
			output::DBGGeneric,
			CLASS,
			L"OnCreateProfile", // STRING_OK
			L"Creating profile \"%ws\"\n", // STRING_OK
			szProfile.c_str());

		EC_H_S(mapi::profile::HrCreateProfile(szProfile));

		// Since we may have created a profile, update even if we failed.
		OnRefreshView(); // Update the view since we don't have notifications here.
	}

	void CProfileListDlg::OnDeleteSelectedItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			// Find the highlighted item AttachNum
			if (!lpListData) break;
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (!contents) break;

			output::DebugPrintEx(
				output::DBGDeleteSelectedItem,
				CLASS,
				L"OnDeleteSelectedItem",
				L"Deleting profile \"%ws\"\n",
				contents->m_szProfileDisplayName.c_str());

			EC_H_S(mapi::profile::HrRemoveProfile(contents->m_szProfileDisplayName));
		}

		OnRefreshView(); // Update the view since we don't have notifications here.
	}

	void CProfileListDlg::OnGetProfileServiceVersion()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			// Find the highlighted item AttachNum
			if (!lpListData) break;
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (!contents) break;

			output::DebugPrintEx(
				output::DBGDeleteSelectedItem,
				CLASS,
				L"OnGetProfileServiceVersion",
				L"Getting profile service version for \"%ws\"\n",
				contents->m_szProfileDisplayName.c_str());

			ULONG ulServerVersion = 0;
			mapi::profile::EXCHANGE_STORE_VERSION_NUM storeVersion = {0};
			auto bFoundServerVersion = false;
			auto bFoundServerFullVersion = false;

			WC_H_S(GetProfileServiceVersion(
				contents->m_szProfileDisplayName,
				&ulServerVersion,
				&storeVersion,
				&bFoundServerVersion,
				&bFoundServerFullVersion));

			editor::CEditor MyData(
				this, IDS_PROFILESERVERVERSIONTITLE, IDS_PROFILESERVERVERSIONPROMPT, CEDITOR_BUTTON_OK);

			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_PROFILESERVERVERSION, true));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_PROFILESERVERVERSIONMAJOR, true));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_PROFILESERVERVERSIONMINOR, true));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(3, IDS_PROFILESERVERVERSIONBUILD, true));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_PROFILESERVERVERSIONMINORBUILD, true));

			if (bFoundServerVersion)
			{
				MyData.SetStringf(0, L"%u = 0x%X", ulServerVersion, ulServerVersion); // STRING_OK
				output::DebugPrint(output::DBGGeneric, L"PR_PROFILE_SERVER_VERSION == 0x%08X\n", ulServerVersion);
			}
			else
			{
				MyData.LoadString(0, IDS_NOTFOUND);
			}

			if (bFoundServerFullVersion)
			{
				output::DebugPrint(output::DBGGeneric, L"PR_PROFILE_SERVER_FULL_VERSION = \n");
				MyData.SetStringf(1, L"%u = 0x%X", storeVersion.wMajorVersion, storeVersion.wMajorVersion); // STRING_OK
				MyData.SetStringf(2, L"%u = 0x%X", storeVersion.wMinorVersion, storeVersion.wMinorVersion); // STRING_OK
				MyData.SetStringf(3, L"%u = 0x%X", storeVersion.wBuild, storeVersion.wBuild); // STRING_OK
				MyData.SetStringf(4, L"%u = 0x%X", storeVersion.wMinorBuild, storeVersion.wMinorBuild); // STRING_OK
				output::DebugPrint(
					output::DBGGeneric,
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

			(void) MyData.DisplayDialog();
		}
	}

	void CProfileListDlg::OnSetDefaultProfile()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		// Find the highlighted item AttachNum
		const auto lpListData = m_lpContentsTableListCtrl->GetFirstSelectedItemData();
		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents)
			{
				output::DebugPrintEx(
					output::DBGGeneric,
					CLASS,
					L"OnSetDefaultProfile",
					L"Setting profile \"%ws\" as default\n",
					contents->m_szProfileDisplayName.c_str());

				EC_H_S(mapi::profile::HrSetDefaultProfile(contents->m_szProfileDisplayName));

				OnRefreshView(); // Update the view since we don't have notifications here.
			}
		}
	}

	void CProfileListDlg::OnOpenProfileByName()
	{
		editor::CEditor MyData(this, IDS_OPENPROFILE, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_OPENPROFILEPROMPT, false));

		if (MyData.DisplayDialog())
		{
			const auto szProfileName = MyData.GetStringW(0);
			if (!szProfileName.empty())
			{
				new CMsgServiceTableDlg(m_lpParent, m_lpMapiObjects, szProfileName);
			}
		}
	}

	void CProfileListDlg::HandleCopy()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(output::DBGGeneric, CLASS, L"HandleCopy", L"\n");
		if (!m_lpContentsTableListCtrl) return;

		// Find the highlighted profile
		const auto lpListData = m_lpContentsTableListCtrl->GetFirstSelectedItemData();
		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents)
			{
				cache::CGlobalCache::getInstance().SetProfileToCopy(contents->m_szProfileDisplayName);
			}
		}
	}

	_Check_return_ bool CProfileListDlg::HandlePaste()
	{
		if (CBaseDialog::HandlePaste()) return true;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(output::DBGGeneric, CLASS, L"HandlePaste", L"\n");

		const auto szOldProfile = cache::CGlobalCache::getInstance().GetProfileToCopy();

		editor::CEditor MyData(this, IDS_COPYPROFILE, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_COPYPROFILEPROMPT, szOldProfile, false));

		if (MyData.DisplayDialog())
		{
			const auto szNewProfile = MyData.GetStringW(0);

			WC_MAPI_S(mapi::profile::HrCopyProfile(szOldProfile, szNewProfile));

			OnRefreshView(); // Update the view since we don't have notifications here.
		}

		return true;
	}

	void CProfileListDlg::OnExportProfile()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(output::DBGGeneric, CLASS, L"OnExportProfile", L"\n");
		if (!m_lpContentsTableListCtrl) return;

		// Find the highlighted profile
		const auto lpListData = m_lpContentsTableListCtrl->GetFirstSelectedItemData();
		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents)
			{
				auto file = file::CFileDialogExW::SaveAs(
					L"xml", // STRING_OK
					contents->m_szProfileDisplayName + L".xml",
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					strings::loadstring(IDS_XMLFILES),
					this);
				if (!file.empty())
				{
					output::ExportProfile(contents->m_szProfileDisplayName, strings::emptystring, false, file);
				}
			}
		}
	}
} // namespace dialog