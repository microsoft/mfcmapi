#include <StdAfx.h>
#include <UI/Dialogs/BaseDialog.h>
#include <UI/FakeSplitter.h>
#include <core/mapi/cache/mapiObjects.h>
#include <UI/ParentWnd.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/HexEditor.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/AboutDlg.h>
#include <core/mapi/adviseSink.h>
#include <core/mapi/extraPropTags.h>
#include <Msi.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/cache/globalCache.h>
#include <UI/Dialogs/ContentsTable/MainDlg.h>
#include <UI/Dialogs/Editors/DbgView.h>
#include <UI/Dialogs/Editors/Options.h>
#include <Windows.h>
#include <malloc.h>
#include <core/mapi/version.h>
#include <UI/addinui.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CBaseDialog";

	CBaseDialog::CBaseDialog(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects, // Pass NULL to create a new m_lpMapiObjects,
		const ULONG ulAddInContext)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_szTitle = strings::loadstring(IDS_BASEDIALOG);

		// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
		m_hIcon = EC_D(HICON, AfxGetApp()->LoadIcon(IDR_MAINFRAME));

		// Let the parent know we have a status bar so we can draw our border correctly
		SetStatusHeight(GetSystemMetrics(SM_CXSIZEFRAME) + ui::GetTextHeight(::GetDesktopWindow()));

		m_lpParent = pParentWnd;
		if (m_lpParent) m_lpParent->AddRef();

		m_ulAddInContext = ulAddInContext;

		m_lpMapiObjects = std::make_shared<cache::CMapiObjects>(lpMapiObjects);
	}

	CBaseDialog::~CBaseDialog()
	{
		TRACE_DESTRUCTOR(CLASS);
		const auto hMenu = ::GetMenu(this->m_hWnd);
		if (hMenu)
		{
			ui::DeleteMenuEntries(hMenu);
			DestroyMenu(hMenu);
		}

		CWnd::DestroyWindow();
		OnNotificationsOff();
		if (m_lpParent) m_lpParent->Release();
	}

	STDMETHODIMP_(ULONG) CBaseDialog::AddRef()
	{
		const auto lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(CLASS, lCount);
		output::DebugPrint(output::dbgLevel::RefCount, L"CBaseDialog::AddRef(\"%ws\")\n", m_szTitle.c_str());
		return lCount;
	}

	STDMETHODIMP_(ULONG) CBaseDialog::Release()
	{
		const auto lCount = InterlockedDecrement(&m_cRef);
		TRACE_RELEASE(CLASS, lCount);
		output::DebugPrint(output::dbgLevel::RefCount, L"CBaseDialog::Release(\"%ws\")\n", m_szTitle.c_str());
		if (!lCount) delete this;
		return lCount;
	}

	BEGIN_MESSAGE_MAP(CBaseDialog, CMyDialog)
	ON_WM_ACTIVATE()
	ON_WM_INITMENU()
	ON_WM_MENUSELECT()
	ON_WM_SIZE()

	ON_COMMAND(ID_OPTIONS, OnOptions)
	ON_COMMAND(ID_OPENMAINWINDOW, OnOpenMainWindow)
	ON_COMMAND(ID_MYHELP, OnHelp)

	ON_COMMAND(ID_NOTIFICATIONSOFF, OnNotificationsOff)
	ON_COMMAND(ID_NOTIFICATIONSON, OnNotificationsOn)
	ON_COMMAND(ID_DISPATCHNOTIFICATIONS, OnDispatchNotifications)

	ON_MESSAGE(WM_MFCMAPI_UPDATESTATUSBAR, msgOnUpdateStatusBar)
	ON_MESSAGE(WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST, msgOnClearSingleMAPIPropList)
	END_MESSAGE_MAP()

	LRESULT CBaseDialog::WindowProc(const UINT message, const WPARAM wParam, const LPARAM lParam)
	{
		switch (message)
		{
		case WM_COMMAND:
		{
			const auto idFrom = LOWORD(wParam);
			// idFrom is the menu item selected
			if (HandleMenu(idFrom)) return S_OK;
			break;
		}
		case WM_PAINT:
			// Paint the status, then let the rest draw itself.
			ui::DrawStatus(
				m_hWnd,
				GetStatusHeight(),
				getStatusMessage(statusPane::data1),
				getStatusWidth(statusPane::data1),
				getStatusMessage(statusPane::data2),
				getStatusWidth(statusPane::data2),
				getStatusMessage(statusPane::infoText));
			break;
		}

		return CMyDialog::WindowProc(message, wParam, lParam);
	}

	BOOL CBaseDialog::OnInitDialog()
	{
		UpdateTitleBarText();

		setStatusWidth(statusPane::data1, 0);
		setStatusWidth(statusPane::data2, 0);
		setStatusWidth(statusPane::infoText, -1);

		SetIcon(m_hIcon, false); // Set small icon - large icon isn't used

		m_lpFakeSplitter.Init(m_hWnd);
		m_lpPropDisplay =
			new controls::sortlistctrl::CSingleMAPIPropListCtrl(&m_lpFakeSplitter, this, m_lpMapiObjects, m_bIsAB);

		if (m_lpPropDisplay) m_lpFakeSplitter.SetPaneTwo(m_lpPropDisplay->GetSafeHwnd());

		return false;
	}

	void CBaseDialog::CreateDialogAndMenu(
		const UINT nIDMenuResource,
		const UINT uiClassMenuResource,
		const UINT uidClassMenuTitle)
	{
		output::DebugPrintEx(
			output::dbgLevel::CreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);

		m_lpszTemplateName = MAKEINTRESOURCE(IDD_BLANK_DIALOG);

		DisplayParentedDialog(nullptr, NULL);

		HMENU hMenu = nullptr;
		if (nIDMenuResource)
		{
			hMenu = ::LoadMenu(nullptr, MAKEINTRESOURCE(nIDMenuResource));
		}
		else
		{
			hMenu = CreateMenu();
		}

		const auto hMenuOld = ::GetMenu(m_hWnd);
		if (hMenuOld) DestroyMenu(hMenuOld);
		::SetMenu(m_hWnd, hMenu);

		AddMenu(hMenu, IDR_MENU_PROPERTY, IDS_PROPERTYMENU, static_cast<unsigned>(-1));

		AddMenu(hMenu, uiClassMenuResource, uidClassMenuTitle, static_cast<unsigned>(-1));

		m_ulAddInMenuItems = ui::addinui::ExtendAddInMenu(hMenu, m_ulAddInContext);

		AddMenu(hMenu, IDR_MENU_TOOLS, IDS_TOOLSMENU, static_cast<unsigned>(-1));

		const auto hSub = GetSubMenu(hMenu, 0);
		::AppendMenu(hSub, MF_SEPARATOR, NULL, nullptr);
		auto szExit = strings::loadstring(IDS_EXIT);
		AppendMenuW(hSub, MF_ENABLED | MF_STRING, IDCANCEL, szExit.c_str());

		// Make sure the menu background is filled in the right color
		auto mi = MENUINFO{};
		mi.cbSize = sizeof(MENUINFO);
		mi.fMask = MIM_BACKGROUND;
		mi.hbrBack = GetSysBrush(ui::uiColor::Background);
		SetMenuInfo(hMenu, &mi);

		ui::ConvertMenuOwnerDraw(hMenu, true);

		// We're done - force our new menu on screen
		DrawMenuBar();
	}

	_Check_return_ bool CBaseDialog::HandleMenu(const WORD wMenuSelect)
	{
		output::DebugPrint(
			output::dbgLevel::Menu, L"CBaseDialog::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
		switch (wMenuSelect)
		{
		case ID_HEXEDITOR:
			OnHexEditor();
			return true;
		case ID_DBGVIEW:
			editor::DisplayDbgView(m_lpParent);
			return true;
		case ID_COMPAREENTRYIDS:
			OnCompareEntryIDs();
			return true;
		case ID_OPENENTRYID:
			OnOpenEntryID(nullptr);
			return true;
		case ID_COMPUTESTOREHASH:
			OnComputeStoreHash();
			return true;
		case ID_COPY:
			HandleCopy();
			return true;
		case ID_PASTE:
			(void) HandlePaste();
			return true;
		case ID_OUTLOOKVERSION:
			OnOutlookVersion();
			return true;
		}
		if (HandleAddInMenu(wMenuSelect)) return true;

		if (m_lpPropDisplay) return m_lpPropDisplay->HandleMenu(wMenuSelect);
		return false;
	}

	void CBaseDialog::OnInitMenu(_In_opt_ CMenu* pMenu)
	{
		const auto bMAPIInitialized = cache::CGlobalCache::getInstance().bMAPIInitialized();

		if (pMenu)
		{
			if (m_lpPropDisplay) m_lpPropDisplay->InitMenu(pMenu);
			pMenu->EnableMenuItem(ID_NOTIFICATIONSON, DIM(bMAPIInitialized && !m_lpBaseAdviseSink));
			pMenu->CheckMenuItem(ID_NOTIFICATIONSON, CHECK(m_lpBaseAdviseSink));
			pMenu->EnableMenuItem(ID_NOTIFICATIONSOFF, DIM(m_lpBaseAdviseSink));
			pMenu->EnableMenuItem(ID_DISPATCHNOTIFICATIONS, DIM(bMAPIInitialized));
		}
		CMyDialog::OnInitMenu(pMenu);
	}

	// Checks flags on add-in menu items to ensure they should be enabled
	// Override to support context sensitive scenarios
	void CBaseDialog::EnableAddInMenus(
		_In_ HMENU hMenu,
		const ULONG ulMenu,
		_In_ LPMENUITEM /*lpAddInMenu*/,
		const UINT uiEnable)
	{
		if (hMenu) EnableMenuItem(hMenu, ulMenu, uiEnable);
	}

	// Help strings can be found in mfcmapi.rc2
	// Will preserve the existing text in the right status pane, restoring it when we stop displaying menus
	void CBaseDialog::OnMenuSelect(const UINT nItemID, const UINT nFlags, HMENU /*hSysMenu*/)
	{
		if (!m_bDisplayingMenuText)
		{
			m_szMenuDisplacedText = getStatusMessage(statusPane::infoText);
		}

		if (nItemID && !(nFlags & (MF_SEPARATOR | MF_POPUP)))
		{
			UpdateStatusBarText(statusPane::infoText, nItemID); // This will LoadString the menu help text for us
			m_bDisplayingMenuText = true;
		}
		else
		{
			m_bDisplayingMenuText = false;
		}
		if (!m_bDisplayingMenuText)
		{
			UpdateStatusBarText(statusPane::infoText, m_szMenuDisplacedText);
		}
	}

	_Check_return_ bool
	CBaseDialog::HandleKeyDown(const UINT nChar, const bool bShift, const bool bCtrl, const bool bMenu)
	{
		output::DebugPrintEx(
			output::dbgLevel::Menu,
			CLASS,
			L"HandleKeyDown",
			L"nChar = 0x%0X, bShift = 0x%X, bCtrl = 0x%X, bMenu = 0x%X\n",
			nChar,
			bShift,
			bCtrl,
			bMenu);
		if (bMenu) return false;

		switch (nChar)
		{
		case 'H':
			if (bCtrl)
			{
				OnHexEditor();
				return true;
			}
			break;
		case 'D':
			if (bCtrl)
			{
				editor::DisplayDbgView(m_lpParent);
				return true;
			}
			break;
		case VK_F1:
			DisplayAboutDlg(this);
			return true;
		case 'S':
			if (bCtrl && m_lpPropDisplay)
			{
				m_lpPropDisplay->SavePropsToXML();
				return true;
			}
			break;
		case VK_DELETE:
			OnDeleteSelectedItem();
			return true;
		case 'X':
			if (bCtrl)
			{
				OnDeleteSelectedItem();
				return true;
			}
			break;
		case 'C':
			if (bCtrl && !bShift)
			{
				HandleCopy();
				return true;
			}
			break;
		case 'V':
			if (bCtrl)
			{
				(void) HandlePaste();
				return true;
			}
			break;
		case 'O':
			if (bCtrl)
			{
				OnOptions();
				return true;
			}
			break;
		case VK_F5:
			if (!bCtrl)
			{
				OnRefreshView();
				return true;
			}
			break;
		case VK_ESCAPE:
			OnEscHit();
			return true;
		case VK_RETURN:
			output::DebugPrint(output::dbgLevel::Menu, L"CBaseDialog::HandleKeyDown posting ID_DISPLAYSELECTEDITEM\n");
			PostMessage(WM_COMMAND, ID_DISPLAYSELECTEDITEM, NULL);
			return true;
		}
		return false;
	}

	// prevent dialog from disappearing on Enter
	void CBaseDialog::OnOK()
	{
		// Now that my controls capture VK_ENTER...this is unneeded...keep it just in case.
	}

	void CBaseDialog::OnCancel()
	{
		ShowWindow(SW_HIDE);
		if (m_lpPropDisplay) m_lpPropDisplay->Release();
		m_lpPropDisplay = nullptr;
		Release();
	}

	void CBaseDialog::OnEscHit()
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnEscHit", L"Not implemented\n");
	}

	void CBaseDialog::OnOptions()
	{
		const bool ulNiceNamesBefore = registry::doColumnNames;
		const bool ulSuppressNotFoundBefore = registry::suppressNotFound;
		const auto bNeedPropRefresh = editor::DisplayOptionsDlg(this);
		const auto bNiceNamesChanged = ulNiceNamesBefore != registry::doColumnNames;
		const auto bSuppressNotFoundChanged = ulSuppressNotFoundBefore != registry::suppressNotFound;
		auto bResetColumns = false;

		if (bNiceNamesChanged || bSuppressNotFoundChanged)
		{
			// We check if this worked so we don't refresh the prop list after resetting the top pane
			// But, if we're a tree view, this won't work at all, so we'll still want to reset props if needed
			bResetColumns = static_cast<bool>(::SendMessage(m_hWnd, WM_MFCMAPI_RESETCOLUMNS, 0, 0));
		}

		if (!bResetColumns && bNeedPropRefresh)
		{
			if (m_lpPropDisplay) WC_H_S(m_lpPropDisplay->RefreshMAPIPropList());
		}
	}

	void CBaseDialog::OnOpenMainWindow()
	{
		auto pMain = new CMainDlg(m_lpParent, m_lpMapiObjects);
		if (pMain) pMain->OnOpenMessageStoreTable();
	}

	void CBaseDialog::HandleCopy() { output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"HandleCopy", L"\n"); }

	_Check_return_ bool CBaseDialog::HandlePaste()
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"HandlePaste", L"\n");
		const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();

		if (m_lpPropDisplay && ulStatus & BUFFER_PROPTAG && ulStatus & BUFFER_SOURCEPROPOBJ)
		{
			m_lpPropDisplay->OnPasteProperty();
			return true;
		}

		return false;
	}

	void CBaseDialog::OnHelp() { DisplayAboutDlg(this); }

	void CBaseDialog::OnDeleteSelectedItem()
	{
		output::DebugPrintEx(
			output::dbgLevel::DeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L" Not Implemented\n");
	}

	void CBaseDialog::OnRefreshView()
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnRefreshView", L" Not Implemented\n");
	}

	void CBaseDialog::OnUpdateSingleMAPIPropListCtrl(
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_opt_ sortlistdata::sortListData* lpListData) const
	{
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"OnUpdateSingleMAPIPropListCtrl", L"Setting item %p\n", lpMAPIProp);

		if (m_lpPropDisplay)
		{
			WC_H_S(m_lpPropDisplay->SetDataSource(lpMAPIProp, lpListData, m_bIsAB));
		}
	}

	void CBaseDialog::AddMenu(HMENU hMenuBar, const UINT uiResource, const UINT uidTitle, const UINT uiPos)
	{
		auto hMenuToAdd = ::LoadMenu(nullptr, MAKEINTRESOURCE(uiResource));

		if (hMenuBar && hMenuToAdd)
		{
			auto szTitle = strings::loadstring(uidTitle);
			InsertMenuW(
				hMenuBar, uiPos, MF_BYPOSITION | MF_POPUP, reinterpret_cast<UINT_PTR>(hMenuToAdd), szTitle.c_str());
			if (IDR_MENU_PROPERTY == uiResource)
			{
				(void) ui::addinui::ExtendAddInMenu(hMenuToAdd, MENU_CONTEXT_PROPERTY);
			}
		}
	}

	void CBaseDialog::OnActivate(const UINT nState, _In_ CWnd* pWndOther, const BOOL bMinimized)
	{
		CMyDialog::OnActivate(nState, pWndOther, bMinimized);
		if (nState == 1 && !bMinimized) EC_B_S(RedrawWindow());
	}

	void CBaseDialog::SetStatusWidths()
	{
		const auto iData1 = !getStatusMessage(statusPane::data1).empty();
		const auto iData2 = !getStatusMessage(statusPane::data2).empty();

		auto sizeData1 = SIZE{};
		auto sizeData2 = SIZE{};
		if (iData1 || iData2)
		{
			const auto hdc = ::GetDC(m_hWnd);
			if (hdc)
			{
				const auto hfontOld = SelectObject(hdc, ui::GetSegoeFontBold());

				if (iData1)
				{
					sizeData1 = ui::GetTextExtentPoint32(hdc, getStatusMessage(statusPane::data1));
				}

				if (iData2)
				{
					sizeData2 = ui::GetTextExtentPoint32(hdc, getStatusMessage(statusPane::data2));
				}

				SelectObject(hdc, hfontOld);
				::ReleaseDC(m_hWnd, hdc);
			}
		}

		const auto nSpacing = GetSystemMetrics(SM_CXEDGE);

		auto iWidthData1 = 0;
		auto iWidthData2 = 0;
		if (sizeData1.cx) iWidthData1 = sizeData1.cx + 4 * nSpacing;
		if (sizeData2.cx) iWidthData2 = sizeData2.cx + 4 * nSpacing;

		setStatusWidth(statusPane::data1, iWidthData1);
		setStatusWidth(statusPane::data2, iWidthData2);
		setStatusWidth(statusPane::infoText, -1);
		auto rcStatus = RECT{};
		::GetClientRect(m_hWnd, &rcStatus);
		rcStatus.top = rcStatus.bottom - GetStatusHeight();

		// Tell the window it needs to paint
		::RedrawWindow(m_hWnd, &rcStatus, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
	}

	void CBaseDialog::OnSize(UINT /* nType*/, const int cx, const int cy)
	{
		const auto hdwp = WC_D(HDWP, BeginDeferWindowPos(1));
		if (hdwp)
		{
			const auto iHeight = GetStatusHeight();
			const auto iNewCY = cy - iHeight;
			auto rcStatus = RECT{};
			::GetClientRect(m_hWnd, &rcStatus);
			if (rcStatus.bottom - rcStatus.top > iHeight)
			{
				rcStatus.top = rcStatus.bottom - iHeight;
			}
			// Tell the status bar it needs repainting
			::InvalidateRect(m_hWnd, &rcStatus, false);

			if (m_lpFakeSplitter && m_lpFakeSplitter.m_hWnd)
			{
				DeferWindowPos(hdwp, m_lpFakeSplitter.m_hWnd, nullptr, 0, 0, cx, iNewCY, SWP_NOZORDER);
			}

			WC_B_S(EndDeferWindowPos(hdwp));
		}
	}

	void CBaseDialog::UpdateStatusBarText(const statusPane nPos, _In_ const std::wstring& szMsg)
	{
		if (nPos < statusPane::numPanes) setStatusMessage(nPos, szMsg);

		SetStatusWidths();
	}

	void __cdecl CBaseDialog::UpdateStatusBarText(const statusPane nPos, const UINT uidMsg)
	{
		UpdateStatusBarText(nPos, uidMsg, strings::emptystring, strings::emptystring, strings::emptystring);
	}

	void __cdecl CBaseDialog::UpdateStatusBarText(const statusPane nPos, UINT const uidMsg, ULONG const ulParam1)
	{
		auto szParam1 = std::to_wstring(ulParam1);
		UpdateStatusBarText(nPos, uidMsg, szParam1, strings::emptystring, strings::emptystring);
	}

	void __cdecl CBaseDialog::UpdateStatusBarText(
		statusPane const nPos,
		UINT uidMsg,
		std::wstring& szParam1,
		std::wstring& szParam2,
		std::wstring& szParam3)
	{
		std::wstring szStatBarString;

		// MAPI Load paths take special handling
		if (uidMsg >= ID_LOADMAPIMENUMIN && uidMsg <= ID_LOADMAPIMENUMAX)
		{
			auto mii = MENUITEMINFOW{};
			mii.cbSize = sizeof(MENUITEMINFO);
			mii.fMask = MIIM_DATA;

			WC_B_S(GetMenuItemInfoW(::GetMenu(m_hWnd), uidMsg, false, &mii));
			if (mii.dwItemData)
			{
				const auto lme = reinterpret_cast<ui::LPMENUENTRY>(mii.dwItemData);
				szStatBarString = strings::formatmessage(IDS_LOADMAPISTATUS, lme->m_pName.c_str());
			}
		}
		else
		{
			const auto lpAddInMenu = ui::addinui::GetAddinMenuItem(m_hWnd, uidMsg);
			if (lpAddInMenu && lpAddInMenu->szHelp)
			{
				szStatBarString = lpAddInMenu->szHelp;
			}
			else
			{
				szStatBarString = strings::formatmessage(uidMsg, szParam1.c_str(), szParam2.c_str(), szParam3.c_str());
			}
		}

		UpdateStatusBarText(nPos, szStatBarString);
	}

	void CBaseDialog::UpdateTitleBarText(_In_ const std::wstring& szMsg) const
	{
		SetTitle(strings::formatmessage(IDS_TITLEBARMESSAGE, m_szTitle.c_str(), szMsg.c_str()));
	}

	void CBaseDialog::UpdateTitleBarText() const
	{
		SetTitle(strings::formatmessage(IDS_TITLEBARPLAIN, m_szTitle.c_str()));
	}

	void CBaseDialog::UpdateStatus(HWND hWndHost, const statusPane pane, const std::wstring& status)
	{
		(void) ::SendMessage(
			hWndHost, WM_MFCMAPI_UPDATESTATUSBAR, static_cast<WPARAM>(pane), reinterpret_cast<LPARAM>(status.c_str()));
	}

	// WM_MFCMAPI_UPDATESTATUSBAR
	_Check_return_ LRESULT CBaseDialog::msgOnUpdateStatusBar(WPARAM wParam, LPARAM const lParam)
	{
		const auto iPane = static_cast<statusPane>(wParam);
		const std::wstring szStr = reinterpret_cast<LPWSTR>(lParam);
		UpdateStatusBarText(iPane, szStr);

		return S_OK;
	}

	// WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST
	_Check_return_ LRESULT CBaseDialog::msgOnClearSingleMAPIPropList(WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);

		return S_OK;
	}

	void CBaseDialog::OnHexEditor() { new editor::CHexEditor(m_lpParent, m_lpMapiObjects); }

	void CBaseDialog::OnOutlookVersion()
	{
		editor::CEditor MyEID(this, IDS_OUTLOOKVERSIONTITLE, NULL, CEDITOR_BUTTON_OK);

		const auto szVersionString = version::GetOutlookVersionString();

		MyEID.AddPane(viewpane::TextPane::CreateMultiLinePane(0, IDS_OUTLOOKVERSIONPROMPT, szVersionString, true));
		(void) MyEID.DisplayDialog();
	}

	void CBaseDialog::OnOpenEntryID(_In_opt_ LPSBinary lpBin)
	{
		if (!m_lpMapiObjects) return;

		editor::CEditor MyEID(this, IDS_OPENEID, IDS_OPENEIDPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyEID.AddPane(
			viewpane::TextPane::CreateSingleLinePane(0, IDS_EID, strings::BinToHexString(lpBin, false), false));

		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		MyEID.AddPane(viewpane::CheckPane::Create(1, IDS_USEMDB, lpMDB != nullptr, lpMDB == nullptr));

		auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
		if (lpAB)
			MyEID.AddPane(viewpane::CheckPane::Create(2, IDS_USEAB, true, lpAB == nullptr));
		else
			MyEID.AddPane(viewpane::CheckPane::Create(2, IDS_USEAB, false, lpAB == nullptr));

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		MyEID.AddPane(viewpane::CheckPane::Create(3, IDS_SESSION, lpMAPISession != nullptr, lpMAPISession == nullptr));

		MyEID.AddPane(viewpane::CheckPane::Create(4, IDS_PASSMAPIMODIFY, false, false));

		MyEID.AddPane(viewpane::CheckPane::Create(5, IDS_PASSMAPINOCACHE, false, false));

		MyEID.AddPane(viewpane::CheckPane::Create(6, IDS_PASSMAPICACHEONLY, false, false));

		MyEID.AddPane(viewpane::CheckPane::Create(7, IDS_EIDBASE64ENCODED, false, false));

		MyEID.AddPane(viewpane::CheckPane::Create(8, IDS_ATTEMPTIADDRBOOKDETAILSCALL, false, lpAB == nullptr));

		MyEID.AddPane(viewpane::CheckPane::Create(9, IDS_EIDISCONTAB, false, false));

		if (!MyEID.DisplayDialog()) return;

		// Get the entry ID as a binary
		const auto bin = MyEID.GetBinary(0, MyEID.GetCheck(7));
		auto sBin = SBinary{static_cast<ULONG>(bin.size()), const_cast<BYTE*>(bin.data())};

		if (MyEID.GetCheck(9))
		{
			(void) mapi::UnwrapContactEntryID(sBin.cb, sBin.lpb, &sBin.cb, &sBin.lpb);
		}

		if (MyEID.GetCheck(8) && lpAB) // Do IAddrBook->Details here
		{
			auto ulUIParam = reinterpret_cast<ULONG_PTR>(static_cast<void*>(m_hWnd));

			EC_H_CANCEL_S(lpAB->Details(
				&ulUIParam,
				nullptr,
				nullptr,
				sBin.cb,
				reinterpret_cast<LPENTRYID>(sBin.lpb),
				nullptr,
				nullptr,
				nullptr,
				DIALOG_MODAL)); // API doesn't like unicode
		}
		else
		{
			ULONG ulObjType = NULL;
			auto lpUnk = mapi::CallOpenEntry<LPUNKNOWN>(
				MyEID.GetCheck(1) ? lpMDB : nullptr,
				MyEID.GetCheck(2) ? lpAB : nullptr,
				nullptr,
				MyEID.GetCheck(3) ? lpMAPISession : nullptr,
				sBin.cb,
				reinterpret_cast<LPENTRYID>(sBin.lpb),
				nullptr,
				(MyEID.GetCheck(4) ? MAPI_MODIFY : MAPI_BEST_ACCESS) | (MyEID.GetCheck(5) ? MAPI_NO_CACHE : 0) |
					(MyEID.GetCheck(6) ? MAPI_CACHE_ONLY : 0),
				&ulObjType);

			if (lpUnk)
			{
				auto szFlags = smartview::InterpretNumberAsStringProp(ulObjType, PR_OBJECT_TYPE);
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"OnOpenEntryID: Got object (%p) of type 0x%08X = %ws\n",
					lpUnk,
					ulObjType,
					szFlags.c_str());

				auto lpTemp = mapi::safe_cast<LPMAPIPROP>(lpUnk);
				if (lpTemp)
				{
					WC_H_S(DisplayObject(lpTemp, ulObjType, objectType::otHierarchy, this));
					lpTemp->Release();
				}

				lpUnk->Release();
			}
		}
	}

	void CBaseDialog::OnCompareEntryIDs()
	{
		if (!m_lpMapiObjects) return;

		auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release

		editor::CEditor MyEIDs(
			this, IDS_COMPAREEIDS, IDS_COMPAREEIDSPROMPTS, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyEIDs.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_EID1, false));
		MyEIDs.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_EID2, false));

		UINT uidDropDown[] = {IDS_DDMESSAGESTORE, IDS_DDSESSION, IDS_DDADDRESSBOOK};
		MyEIDs.AddPane(
			viewpane::DropDownPane::Create(2, IDS_OBJECTFORCOMPAREEID, _countof(uidDropDown), uidDropDown, true));

		MyEIDs.AddPane(viewpane::CheckPane::Create(3, IDS_EIDBASE64ENCODED, false, false));

		if (!MyEIDs.DisplayDialog()) return;

		if (0 == MyEIDs.GetDropDown(2) && !lpMDB || 1 == MyEIDs.GetDropDown(2) && !lpMAPISession ||
			2 == MyEIDs.GetDropDown(2) && !lpAB)
		{
			error::ErrDialog(__FILE__, __LINE__, IDS_EDCOMPAREEID);
			return;
		}

		// Get the entry IDs as a binary
		const auto bin1 = MyEIDs.GetBinary(0, MyEIDs.GetCheck(3));
		auto sBin1 = SBinary{static_cast<ULONG>(bin1.size()), const_cast<BYTE*>(bin1.data())};

		const auto bin2 = MyEIDs.GetBinary(1, MyEIDs.GetCheck(3));
		auto sBin2 = SBinary{static_cast<ULONG>(bin2.size()), const_cast<BYTE*>(bin2.data())};

		auto hRes = S_OK;
		ULONG ulResult = NULL;
		switch (MyEIDs.GetDropDown(2))
		{
		case 0: // Message Store
			hRes = EC_MAPI(lpMDB->CompareEntryIDs(
				sBin1.cb,
				reinterpret_cast<LPENTRYID>(sBin1.lpb),
				sBin2.cb,
				reinterpret_cast<LPENTRYID>(sBin2.lpb),
				NULL,
				&ulResult));
			break;
		case 1: // Session
			hRes = EC_MAPI(lpMAPISession->CompareEntryIDs(
				sBin1.cb,
				reinterpret_cast<LPENTRYID>(sBin1.lpb),
				sBin2.cb,
				reinterpret_cast<LPENTRYID>(sBin2.lpb),
				NULL,
				&ulResult));
			break;
		case 2: // Address Book
			hRes = EC_MAPI(lpAB->CompareEntryIDs(
				sBin1.cb,
				reinterpret_cast<LPENTRYID>(sBin1.lpb),
				sBin2.cb,
				reinterpret_cast<LPENTRYID>(sBin2.lpb),
				NULL,
				&ulResult));
			break;
		}

		if (SUCCEEDED(hRes))
		{
			auto szResult = strings::loadstring(ulResult ? IDS_TRUE : IDS_FALSE);
			const auto szRet = strings::formatmessage(IDS_COMPAREEIDBOOL, ulResult, szResult.c_str());

			editor::CEditor Result(this, IDS_COMPAREEIDSRESULT, NULL, CEDITOR_BUTTON_OK);
			Result.SetPromptPostFix(szRet);
			(void) Result.DisplayDialog();
		}
	}

	void CBaseDialog::OnComputeStoreHash()
	{
		editor::CEditor MyStoreEID(
			this, IDS_COMPUTESTOREHASH, IDS_COMPUTESTOREHASHPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyStoreEID.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_STOREEID, false));
		MyStoreEID.AddPane(viewpane::CheckPane::Create(1, IDS_EIDBASE64ENCODED, false, false));
		MyStoreEID.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_FILENAME, false));
		MyStoreEID.AddPane(viewpane::CheckPane::Create(3, IDS_PUBLICFOLDERSTORE, false, false));

		if (!MyStoreEID.DisplayDialog()) return;

		// Get the entry ID as a binary
		const auto bin = MyStoreEID.GetBinary(0, MyStoreEID.GetCheck(1));
		auto sBin = SBinary{static_cast<ULONG>(bin.size()), const_cast<BYTE*>(bin.data())};

		const auto dwHash = mapi::ComputeStoreHash(
			sBin.cb, sBin.lpb, nullptr, MyStoreEID.GetStringW(2).c_str(), MyStoreEID.GetCheck(3));
		const auto szHash = strings::formatmessage(IDS_STOREHASHVAL, dwHash);

		editor::CEditor Result(this, IDS_STOREHASH, NULL, CEDITOR_BUTTON_OK);
		Result.SetPromptPostFix(szHash);
		(void) Result.DisplayDialog();
	}

	void CBaseDialog::OnNotificationsOn()
	{
		if (m_lpBaseAdviseSink || !m_lpMapiObjects) return;

		auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release

		editor::CEditor MyData(
			this, IDS_NOTIFICATIONS, IDS_NOTIFICATIONSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(flags::AllFlagsToString(flagNotifEventType, true));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_EID, false));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_ULEVENTMASK, false));
		MyData.SetHex(1, fnevNewMail);
		UINT uidDropDown[] = {IDS_DDMESSAGESTORE, IDS_DDSESSION, IDS_DDADDRESSBOOK};
		MyData.AddPane(
			viewpane::DropDownPane::Create(2, IDS_OBJECTFORADVISE, _countof(uidDropDown), uidDropDown, true));

		if (MyData.DisplayDialog())
		{
			if (0 == MyData.GetDropDown(2) && !lpMDB || 1 == MyData.GetDropDown(2) && !lpMAPISession ||
				2 == MyData.GetDropDown(2) && !lpAB)
			{
				error::ErrDialog(__FILE__, __LINE__, IDS_EDADVISE);
				return;
			}

			const auto bin = MyData.GetBinary(0, false);
			auto sBin = SBinary{static_cast<ULONG>(bin.size()), const_cast<BYTE*>(bin.data())};
			// don't actually care if the returning lpEntryID is NULL - Advise can work with that

			m_lpBaseAdviseSink = new (std::nothrow) mapi::adviseSink(m_hWnd, nullptr);

			auto hRes = S_OK;
			if (m_lpBaseAdviseSink)
			{
				switch (MyData.GetDropDown(2))
				{
				case 0:
					hRes = EC_MAPI(lpMDB->Advise(
						sBin.cb,
						reinterpret_cast<LPENTRYID>(sBin.lpb),
						MyData.GetHex(1),
						static_cast<IMAPIAdviseSink*>(m_lpBaseAdviseSink),
						&m_ulBaseAdviseConnection));
					m_lpBaseAdviseSink->SetAdviseTarget(lpMDB);
					m_ulBaseAdviseObjectType = MAPI_STORE;
					break;
				case 1:
					hRes = EC_MAPI(lpMAPISession->Advise(
						sBin.cb,
						reinterpret_cast<LPENTRYID>(sBin.lpb),
						MyData.GetHex(1),
						static_cast<IMAPIAdviseSink*>(m_lpBaseAdviseSink),
						&m_ulBaseAdviseConnection));
					m_ulBaseAdviseObjectType = MAPI_SESSION;
					break;
				case 2:
					hRes = EC_MAPI(lpAB->Advise(
						sBin.cb,
						reinterpret_cast<LPENTRYID>(sBin.lpb),
						MyData.GetHex(1),
						static_cast<IMAPIAdviseSink*>(m_lpBaseAdviseSink),
						&m_ulBaseAdviseConnection));
					m_lpBaseAdviseSink->SetAdviseTarget(lpAB);
					m_ulBaseAdviseObjectType = MAPI_ADDRBOOK;
					break;
				}

				if (SUCCEEDED(hRes))
				{
					if (0 == MyData.GetDropDown(2) && lpMDB)
					{
						mapi::ForceRop(lpMDB);
					}
				}
				else // if we failed to advise, then we don't need the advise sink object
				{
					if (m_lpBaseAdviseSink) m_lpBaseAdviseSink->Release();
					m_lpBaseAdviseSink = nullptr;
					m_ulBaseAdviseObjectType = NULL;
					m_ulBaseAdviseConnection = NULL;
				}
			}
		}
	}

	void CBaseDialog::OnNotificationsOff()
	{
		if (m_ulBaseAdviseConnection && m_lpMapiObjects)
		{
			switch (m_ulBaseAdviseObjectType)
			{
			case MAPI_SESSION:
			{
				auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
				if (lpMAPISession) EC_MAPI_S(lpMAPISession->Unadvise(m_ulBaseAdviseConnection));
				break;
			}
			case MAPI_STORE:
			{
				auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				if (lpMDB) EC_MAPI_S(lpMDB->Unadvise(m_ulBaseAdviseConnection));
				break;
			}
			case MAPI_ADDRBOOK:
			{
				auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
				if (lpAB) EC_MAPI_S(lpAB->Unadvise(m_ulBaseAdviseConnection));
				break;
			}
			}
		}

		if (m_lpBaseAdviseSink) m_lpBaseAdviseSink->Release();
		m_lpBaseAdviseSink = nullptr;
		m_ulBaseAdviseObjectType = NULL;
		m_ulBaseAdviseConnection = NULL;
	}

	void CBaseDialog::OnDispatchNotifications() { EC_MAPI_S(HrDispatchNotifications(NULL)); }

	_Check_return_ bool CBaseDialog::HandleAddInMenu(const WORD wMenuSelect)
	{
		output::DebugPrintEx(
			output::dbgLevel::AddInPlumbing, CLASS, L"HandleAddInMenu", L"wMenuSelect = 0x%08X\n", wMenuSelect);
		return false;
	}

	_Check_return_ ui::CParentWnd* CBaseDialog::GetParentWnd() const { return m_lpParent; }

	_Check_return_ std::shared_ptr<cache::CMapiObjects> CBaseDialog::GetMapiObjects() const { return m_lpMapiObjects; }
} // namespace dialog