// Displays the list of providers in a message service in a profile
#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/ProviderTableDlg.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/columnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <core/mapi/mapiProfileFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <core/sortlistdata/contentsData.h>
#include <UI/addinui.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CProviderTableDlg";

	CProviderTableDlg::CProviderTableDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_ LPMAPITABLE lpMAPITable,
		_In_ LPPROVIDERADMIN lpProviderAdmin)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_PROVIDERS,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  lpMAPITable,
			  &columns::sptPROVIDERCols.tags,
			  columns::PROVIDERColumns,
			  NULL,
			  MENU_CONTEXT_PROFILE_PROVIDERS)
	{
		TRACE_CONSTRUCTOR(CLASS);

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_PROVIDER);

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

	_Check_return_ LPMAPIPROP CProviderTableDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum /*bModify*/)
	{
		if (!m_lpContentsTableListCtrl || !m_lpProviderAdmin) return nullptr;
		output::DebugPrintEx(output::DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		LPPROFSECT lpProfSect = nullptr;
		const auto lpListData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);
		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents)
			{
				const auto lpProviderUID = contents->m_lpProviderUID;
				if (lpProviderUID)
				{
					lpProfSect = mapi::profile::OpenProfileSection(m_lpProviderAdmin, lpProviderUID);
				}
			}
		}

		return lpProfSect;
	}

	void CProviderTableDlg::OnOpenProfileSection()
	{
		if (!m_lpProviderAdmin) return;

		editor::CEditor MyUID(
			this, IDS_OPENPROFSECT, IDS_OPENPROFSECTPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyUID.AddPane(viewpane::DropDownPane::CreateGuid(0, IDS_MAPIUID, false));
		MyUID.AddPane(viewpane::CheckPane::Create(1, IDS_MAPIUIDBYTESWAPPED, false, false));

		if (!MyUID.DisplayDialog()) return;

		auto guid = MyUID.GetSelectedGUID(0, MyUID.GetCheck(1));
		auto MapiUID = SBinary{sizeof(GUID), reinterpret_cast<LPBYTE>(&guid)};

		auto lpProfSect = mapi::profile::OpenProfileSection(m_lpProviderAdmin, &MapiUID);
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

	void CProviderTableDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER /*lpContainer*/)
	{
		if (lpParams)
		{
			lpParams->lpProfSect = dynamic_cast<LPPROFSECT>(lpMAPIProp);
		}

		ui::addinui::InvokeAddInMenu(lpParams);
	}
} // namespace dialog