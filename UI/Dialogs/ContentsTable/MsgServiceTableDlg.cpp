// Displays the list of services in a profile
#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/MsgServiceTableDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/Cache/MapiObjects.h>
#include <MAPI/ColumnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/ContentsTable/ProviderTableDlg.h>
#include <MAPI/MAPIProfileFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Controls/SortList/ContentsData.h>
#include <MAPI/MAPIFunctions.h>
#include <UI/addinui.h>
#include <IO/MFCOutput.h>
#include <IO/Registry.h>
#include <core/utility/strings.h>

namespace dialog
{
	static std::wstring CLASS = L"CMsgServiceTableDlg";

	CMsgServiceTableDlg::CMsgServiceTableDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_ const std::string& szProfileName)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_SERVICES,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  nullptr,
			  &columns::sptSERVICECols.tags,
			  columns::SERVICEColumns,
			  IDR_MENU_MSGSERVICE_POPUP,
			  MENU_CONTEXT_PROFILE_SERVICES)
	{
		TRACE_CONSTRUCTOR(CLASS);

		m_lpServiceAdmin = nullptr;

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_MSGSERVICE);

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
				const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
				pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
				pMenu->EnableMenuItem(ID_CONFIGUREMSGSERVICE, DIMMSOK(iNumSel));
			}
		}

		CContentsTableDlg::OnInitMenu(pMenu);
	}

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	void CMsgServiceTableDlg::OnRefreshView()
	{
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L"\n");

		// Make sure we've got something to work with
		if (m_szProfileName.empty() || !m_lpContentsTableListCtrl || !m_lpMapiObjects) return;

		// cancel any loading which may be occuring
		if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();

		// Clean up our table and admin in reverse order from which we obtained them
		// Failure to do this leads to crashes in Outlook's profile code
		m_lpContentsTableListCtrl->SetContentsTable(nullptr, dfNormal, NULL);

		if (m_lpServiceAdmin) m_lpServiceAdmin->Release();
		m_lpServiceAdmin = nullptr;

		LPPROFADMIN lpProfAdmin = nullptr;
		EC_MAPI_S(MAPIAdminProfiles(0, &lpProfAdmin));

		if (lpProfAdmin)
		{
			EC_MAPI_S(lpProfAdmin->AdminServices(
				reinterpret_cast<LPTSTR>(const_cast<LPSTR>(m_szProfileName.c_str())),
				reinterpret_cast<LPTSTR>(""),
				NULL,
				MAPI_DIALOG,
				&m_lpServiceAdmin));
			if (m_lpServiceAdmin)
			{
				LPMAPITABLE lpServiceTable = nullptr;

				EC_MAPI_S(m_lpServiceAdmin->GetMsgServiceTable(
					0, // fMapiUnicode is not supported
					&lpServiceTable));

				if (lpServiceTable)
				{
					m_lpContentsTableListCtrl->SetContentsTable(lpServiceTable, dfNormal, NULL);

					lpServiceTable->Release();
				}
			}

			lpProfAdmin->Release();
		}
	}

	void CMsgServiceTableDlg::OnDisplayItem()
	{
		LPPROVIDERADMIN lpProviderAdmin = nullptr;
		LPMAPITABLE lpProviderTable = nullptr;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl || !m_lpServiceAdmin) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (lpListData && lpListData->Contents())
			{
				const auto lpServiceUID = lpListData->Contents()->m_lpServiceUID;
				if (lpServiceUID)
				{
					EC_MAPI_S(m_lpServiceAdmin->AdminProviders(
						reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb),
						0, // fMapiUnicode is not supported
						&lpProviderAdmin));

					if (lpProviderAdmin)
					{
						EC_MAPI_S(lpProviderAdmin->GetProviderTable(
							0, // fMapiUnicode is not supported
							&lpProviderTable));

						if (lpProviderTable)
						{
							new CProviderTableDlg(m_lpParent, m_lpMapiObjects, lpProviderTable, lpProviderAdmin);
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
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl || !m_lpServiceAdmin) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (lpListData && lpListData->Contents())
			{
				const auto lpServiceUID = lpListData->Contents()->m_lpServiceUID;
				if (lpServiceUID)
				{
					EC_H_CANCEL_S(m_lpServiceAdmin->ConfigureMsgService(
						reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb),
						reinterpret_cast<ULONG_PTR>(m_hWnd),
						SERVICE_UI_ALWAYS,
						0,
						nullptr));
				}
			}
		}
	}

	_Check_return_ LPMAPIPROP CMsgServiceTableDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum /*bModify*/)
	{
		if (!m_lpServiceAdmin || !m_lpContentsTableListCtrl) return nullptr;

		output::DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		LPPROFSECT lpProfSect = nullptr;
		const auto lpListData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);
		if (lpListData && lpListData->Contents())
		{
			const auto lpServiceUID = lpListData->Contents()->m_lpServiceUID;
			if (lpServiceUID)
			{
				lpProfSect = mapi::profile::OpenProfileSection(m_lpServiceAdmin, lpServiceUID);
			}
		}

		return lpProfSect;
	}

	void CMsgServiceTableDlg::OnOpenProfileSection()
	{
		if (!m_lpServiceAdmin) return;

		editor::CEditor MyUID(
			this, IDS_OPENPROFSECT, IDS_OPENPROFSECTPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyUID.AddPane(viewpane::DropDownPane::CreateGuid(0, IDS_MAPIUID, false));
		MyUID.AddPane(viewpane::CheckPane::Create(1, IDS_MAPIUIDBYTESWAPPED, false, false));

		if (!MyUID.DisplayDialog()) return;

		auto guid = MyUID.GetSelectedGUID(0, MyUID.GetCheck(1));
		auto MapiUID = SBinary{sizeof(GUID), reinterpret_cast<LPBYTE>(&guid)};

		auto lpProfSect = mapi::profile::OpenProfileSection(m_lpServiceAdmin, &MapiUID);
		if (lpProfSect)
		{
			auto lpTemp = mapi::safe_cast<LPMAPIPROP>(lpProfSect);
			if (lpTemp)
			{
				EC_H_S(DisplayObject(lpTemp, MAPI_PROFSECT, otContents, this));
				lpTemp->Release();
			}

			lpProfSect->Release();
		}

		MAPIFreeBuffer(MapiUID.lpb);
	}

	void CMsgServiceTableDlg::OnDeleteSelectedItem()
	{
		if (!m_lpServiceAdmin || !m_lpContentsTableListCtrl) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			// Find the highlighted item AttachNum
			if (!lpListData || !lpListData->Contents()) break;

			output::DebugPrintEx(
				DBGDeleteSelectedItem,
				CLASS,
				L"OnDeleteSelectedItem",
				L"Deleting service from \"%hs\"\n",
				lpListData->Contents()->m_szProfileDisplayName.c_str());

			const auto lpServiceUID = lpListData->Contents()->m_lpServiceUID;
			if (lpServiceUID)
			{
				WC_MAPI_S(m_lpServiceAdmin->DeleteMsgService(reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb)));
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
			lpParams->lpProfSect = mapi::safe_cast<LPPROFSECT>(lpMAPIProp);
		}

		ui::addinui::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpProfSect)
		{
			lpParams->lpProfSect->Release();
			lpParams->lpProfSect = nullptr;
		}
	}
} // namespace dialog