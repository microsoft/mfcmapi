#include <StdAfx.h>
#include <UI/Dialogs/BaseDialog.h>
#include <UI/FakeSplitter.h>
#include <MAPI/Cache/MapiObjects.h>
#include <UI/ParentWnd.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/HexEditor.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/UIFunctions.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/InterpretProp.h>
#include <UI/Dialogs/AboutDlg.h>
#include <MAPI/AdviseSink.h>
#include <Interpret/ExtraPropTags.h>
#include <Msi.h>
#include <ImportProcs.h>
#include <Interpret/SmartView/SmartView.h>
#include <MAPI/Cache/GlobalCache.h>
#include <UI/Dialogs/ContentsTable/MainDlg.h>
#include <UI/Dialogs/Editors/DbgView.h>
#include <UI/Dialogs/Editors/Options.h>

namespace dialog
{
	static std::wstring CLASS = L"CBaseDialog";

	CBaseDialog::CBaseDialog(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects, // Pass NULL to create a new m_lpMapiObjects,
		ULONG ulAddInContext)
	{
		TRACE_CONSTRUCTOR(CLASS);
		auto hRes = S_OK;
		m_szTitle = strings::loadstring(IDS_BASEDIALOG);
		m_bDisplayingMenuText = false;

		m_lpBaseAdviseSink = nullptr;
		m_ulBaseAdviseConnection = NULL;
		m_ulBaseAdviseObjectType = NULL;

		m_bIsAB = false;

		// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
		EC_D(m_hIcon, AfxGetApp()->LoadIcon(IDR_MAINFRAME));

		m_cRef = 1;
		m_lpPropDisplay = nullptr;
		m_lpFakeSplitter = nullptr;
		// Let the parent know we have a status bar so we can draw our border correctly
		SetStatusHeight(GetSystemMetrics(SM_CXSIZEFRAME) + ui::GetTextHeight(::GetDesktopWindow()));

		m_lpParent = pParentWnd;
		if (m_lpParent) m_lpParent->AddRef();

		m_lpContainer = nullptr;
		m_ulAddInContext = ulAddInContext;
		m_ulAddInMenuItems = NULL;

		m_lpMapiObjects = new cache::CMapiObjects(lpMapiObjects);
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
		if (m_lpContainer) m_lpContainer->Release();
		if (m_lpMapiObjects) m_lpMapiObjects->Release();
		if (m_lpParent) m_lpParent->Release();
	}

	STDMETHODIMP_(ULONG) CBaseDialog::AddRef()
	{
		const auto lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(CLASS, lCount);
		output::DebugPrint(DBGRefCount, L"CBaseDialog::AddRef(\"%ws\")\n", m_szTitle.c_str());
		return lCount;
	}

	STDMETHODIMP_(ULONG) CBaseDialog::Release()
	{
		const auto lCount = InterlockedDecrement(&m_cRef);
		TRACE_RELEASE(CLASS, lCount);
		output::DebugPrint(DBGRefCount, L"CBaseDialog::Release(\"%ws\")\n", m_szTitle.c_str());
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

	LRESULT CBaseDialog::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
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
				m_StatusMessages[STATUSDATA1],
				m_StatusWidth[STATUSDATA1],
				m_StatusMessages[STATUSDATA2],
				m_StatusWidth[STATUSDATA2],
				m_StatusMessages[STATUSINFOTEXT]);
			break;
		}

		return CMyDialog::WindowProc(message, wParam, lParam);
	}

	BOOL CBaseDialog::OnInitDialog()
	{
		UpdateTitleBarText();

		m_StatusWidth[STATUSDATA1] = 0;
		m_StatusWidth[STATUSDATA2] = 0;
		m_StatusWidth[STATUSINFOTEXT] = -1;

		SetIcon(m_hIcon, false); // Set small icon - large icon isn't used

		m_lpFakeSplitter = new controls::CFakeSplitter(this);

		if (m_lpFakeSplitter)
		{
			m_lpPropDisplay = new controls::sortlistctrl::CSingleMAPIPropListCtrl(m_lpFakeSplitter, this, m_lpMapiObjects, m_bIsAB);

			if (m_lpPropDisplay)
				m_lpFakeSplitter->SetPaneTwo(m_lpPropDisplay);
		}
		return false;
	}

	void CBaseDialog::CreateDialogAndMenu(UINT nIDMenuResource, UINT uiClassMenuResource, UINT uidClassMenuTitle)
	{
		output::DebugPrintEx(DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);

		m_lpszTemplateName = MAKEINTRESOURCE(IDD_BLANK_DIALOG);

		DisplayParentedDialog(nullptr, NULL);

		HMENU hMenu = nullptr;
		if (nIDMenuResource)
		{
			hMenu = ::LoadMenu(nullptr, MAKEINTRESOURCE(nIDMenuResource));
		}
		else
		{
			hMenu = ::CreateMenu();
		}

		const auto hMenuOld = ::GetMenu(m_hWnd);
		if (hMenuOld) ::DestroyMenu(hMenuOld);
		::SetMenu(m_hWnd, hMenu);

		AddMenu(hMenu, IDR_MENU_PROPERTY, IDS_PROPERTYMENU, static_cast<unsigned>(-1));

		AddMenu(hMenu, uiClassMenuResource, uidClassMenuTitle, static_cast<unsigned>(-1));

		m_ulAddInMenuItems = addin::ExtendAddInMenu(hMenu, m_ulAddInContext);

		AddMenu(hMenu, IDR_MENU_TOOLS, IDS_TOOLSMENU, static_cast<unsigned>(-1));

		const auto hSub = ::GetSubMenu(hMenu, 0);
		::AppendMenu(hSub, MF_SEPARATOR, NULL, nullptr);
		auto szExit = strings::loadstring(IDS_EXIT);
		::AppendMenuW(hSub, MF_ENABLED | MF_STRING, IDCANCEL, szExit.c_str());

		// Make sure the menu background is filled in the right color
		MENUINFO mi = { 0 };
		mi.cbSize = sizeof(MENUINFO);
		mi.fMask = MIM_BACKGROUND;
		mi.hbrBack = ui::GetSysBrush(ui::cBackground);
		::SetMenuInfo(hMenu, &mi);

		ui::ConvertMenuOwnerDraw(hMenu, true);

		// We're done - force our new menu on screen
		DrawMenuBar();
	}

	_Check_return_ bool CBaseDialog::HandleMenu(WORD wMenuSelect)
	{
		output::DebugPrint(DBGMenu, L"CBaseDialog::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
		switch (wMenuSelect)
		{
		case ID_HEXEDITOR: OnHexEditor(); return true;
		case ID_DBGVIEW: editor::DisplayDbgView(m_lpParent); return true;
		case ID_COMPAREENTRYIDS: OnCompareEntryIDs(); return true;
		case ID_OPENENTRYID: OnOpenEntryID(nullptr); return true;
		case ID_COMPUTESTOREHASH: OnComputeStoreHash(); return true;
		case ID_COPY: HandleCopy(); return true;
		case ID_PASTE: (void)HandlePaste(); return true;
		case ID_OUTLOOKVERSION: OnOutlookVersion(); return true;
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
	void CBaseDialog::EnableAddInMenus(_In_ HMENU hMenu, ULONG ulMenu, _In_ LPMENUITEM /*lpAddInMenu*/, UINT uiEnable)
	{
		if (hMenu) ::EnableMenuItem(hMenu, ulMenu, uiEnable);
	}

	// Help strings can be found in mfcmapi.rc2
	// Will preserve the existing text in the right status pane, restoring it when we stop displaying menus
	void CBaseDialog::OnMenuSelect(UINT nItemID, UINT nFlags, HMENU /*hSysMenu*/)
	{
		if (!m_bDisplayingMenuText)
		{
			m_szMenuDisplacedText = m_StatusMessages[STATUSINFOTEXT];
		}

		if (nItemID && !(nFlags & (MF_SEPARATOR | MF_POPUP)))
		{
			UpdateStatusBarText(STATUSINFOTEXT, nItemID); // This will LoadString the menu help text for us
			m_bDisplayingMenuText = true;
		}
		else
		{
			m_bDisplayingMenuText = false;
		}
		if (!m_bDisplayingMenuText)
		{
			UpdateStatusBarText(STATUSINFOTEXT, m_szMenuDisplacedText);
		}
	}

	_Check_return_ bool CBaseDialog::HandleKeyDown(UINT nChar, bool bShift, bool bCtrl, bool bMenu)
	{
		output::DebugPrintEx(DBGMenu, CLASS, L"HandleKeyDown", L"nChar = 0x%0X, bShift = 0x%X, bCtrl = 0x%X, bMenu = 0x%X\n",
			nChar, bShift, bCtrl, bMenu);
		if (bMenu) return false;

		switch (nChar)
		{
		case 'H':
			if (bCtrl)
			{
				OnHexEditor(); return true;
			}
			break;
		case 'D':
			if (bCtrl)
			{
				editor::DisplayDbgView(m_lpParent); return true;
			}
			break;
		case VK_F1:
			DisplayAboutDlg(this); return true;
		case 'S':
			if (bCtrl && m_lpPropDisplay)
			{
				m_lpPropDisplay->SavePropsToXML(); return true;
			}
			break;
		case VK_DELETE:
			OnDeleteSelectedItem();
			return true;
		case 'X':
			if (bCtrl)
			{
				OnDeleteSelectedItem(); return true;
			}
			break;
		case 'C':
			if (bCtrl && !bShift)
			{
				HandleCopy(); return true;
			}
			break;
		case 'V':
			if (bCtrl)
			{
				(void)HandlePaste(); return true;
			}
			break;
		case 'O':
			if (bCtrl)
			{
				OnOptions(); return true;
			}
			break;
		case VK_F5:
			if (!bCtrl)
			{
				OnRefreshView(); return true;
			}
			break;
		case VK_ESCAPE:
			OnEscHit();
			return true;
		case VK_RETURN:
			output::DebugPrint(DBGMenu, L"CBaseDialog::HandleKeyDown posting ID_DISPLAYSELECTEDITEM\n");
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
		delete m_lpFakeSplitter;
		m_lpFakeSplitter = nullptr;
		Release();
	}

	void CBaseDialog::OnEscHit()
	{
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnEscHit", L"Not implemented\n");
	}

	void CBaseDialog::OnOptions()
	{
		const auto ulNiceNamesBefore = registry::RegKeys[registry::regkeyDO_COLUMN_NAMES].ulCurDWORD;
		const auto ulSuppressNotFoundBefore = registry::RegKeys[registry::regkeySUPPRESS_NOT_FOUND].ulCurDWORD;
		const auto bNeedPropRefresh = editor::DisplayOptionsDlg(this);
		const auto bNiceNamesChanged = ulNiceNamesBefore != registry::RegKeys[registry::regkeyDO_COLUMN_NAMES].ulCurDWORD;
		const auto bSuppressNotFoundChanged = ulSuppressNotFoundBefore != registry::RegKeys[registry::regkeySUPPRESS_NOT_FOUND].ulCurDWORD;
		auto hRes = S_OK;
		auto bResetColumns = false;

		if (bNiceNamesChanged || bSuppressNotFoundChanged)
		{
			// We check if this worked so we don't refresh the prop list after resetting the top pane
			// But, if we're a tree view, this won't work at all, so we'll still want to reset props if needed
			bResetColumns = static_cast<bool>(::SendMessage(m_hWnd, WM_MFCMAPI_RESETCOLUMNS, 0, 0));
		}

		if (!bResetColumns && bNeedPropRefresh)
		{
			if (m_lpPropDisplay) WC_H(m_lpPropDisplay->RefreshMAPIPropList());
		}
	}

	void CBaseDialog::OnOpenMainWindow()
	{
		auto pMain = new CMainDlg(m_lpParent, m_lpMapiObjects);
		if (pMain) pMain->OnOpenMessageStoreTable();
	}

	void CBaseDialog::HandleCopy()
	{
		output::DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
	}

	_Check_return_ bool CBaseDialog::HandlePaste()
	{
		output::DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");
		const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();

		if (m_lpPropDisplay && ulStatus & BUFFER_PROPTAG && ulStatus & BUFFER_SOURCEPROPOBJ)
		{
			m_lpPropDisplay->OnPasteProperty();
			return true;
		}

		return false;
	}

	void CBaseDialog::OnHelp()
	{
		DisplayAboutDlg(this);
	}

	void CBaseDialog::OnDeleteSelectedItem()
	{
		output::DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L" Not Implemented\n");
	}

	void CBaseDialog::OnRefreshView()
	{
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L" Not Implemented\n");
	}

	void CBaseDialog::OnUpdateSingleMAPIPropListCtrl(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ controls::sortlistdata::SortListData* lpListData) const
	{
		auto hRes = S_OK;
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnUpdateSingleMAPIPropListCtrl", L"Setting item %p\n", lpMAPIProp);

		if (m_lpPropDisplay)
		{
			WC_H(m_lpPropDisplay->SetDataSource(
				lpMAPIProp,
				lpListData,
				m_bIsAB));
		}
	}

	void CBaseDialog::AddMenu(HMENU hMenuBar, UINT uiResource, UINT uidTitle, UINT uiPos)
	{
		auto hMenuToAdd = ::LoadMenu(nullptr, MAKEINTRESOURCE(uiResource));

		if (hMenuBar && hMenuToAdd)
		{
			auto szTitle = strings::loadstring(uidTitle);
			::InsertMenuW(hMenuBar, uiPos, MF_BYPOSITION | MF_POPUP, reinterpret_cast<UINT_PTR>(hMenuToAdd), szTitle.c_str());
			if (IDR_MENU_PROPERTY == uiResource)
			{
				(void)addin::ExtendAddInMenu(hMenuToAdd, MENU_CONTEXT_PROPERTY);
			}
		}
	}

	void CBaseDialog::OnActivate(UINT nState, _In_ CWnd* pWndOther, BOOL bMinimized)
	{
		auto hRes = S_OK;
		CMyDialog::OnActivate(nState, pWndOther, bMinimized);
		if (nState == 1 && !bMinimized) EC_B(RedrawWindow());
	}

	void CBaseDialog::SetStatusWidths()
	{
		const auto iData1 = !m_StatusMessages[STATUSDATA1].empty();
		const auto iData2 = !m_StatusMessages[STATUSDATA2].empty();

		SIZE sizeData1 = { 0 };
		SIZE sizeData2 = { 0 };
		if (iData1 || iData2)
		{
			const auto hdc = ::GetDC(m_hWnd);
			if (hdc)
			{
				const auto hfontOld = ::SelectObject(hdc, ui::GetSegoeFontBold());

				if (iData1)
				{
					sizeData1 = ui::GetTextExtentPoint32(hdc, m_StatusMessages[STATUSDATA1]);
				}

				if (iData2)
				{
					sizeData2 = ui::GetTextExtentPoint32(hdc, m_StatusMessages[STATUSDATA2]);
				}

				::SelectObject(hdc, hfontOld);
				::ReleaseDC(m_hWnd, hdc);
			}
		}

		const auto nSpacing = GetSystemMetrics(SM_CXEDGE);

		auto iWidthData1 = 0;
		auto iWidthData2 = 0;
		if (sizeData1.cx) iWidthData1 = sizeData1.cx + 4 * nSpacing;
		if (sizeData2.cx) iWidthData2 = sizeData2.cx + 4 * nSpacing;

		m_StatusWidth[STATUSDATA1] = iWidthData1;
		m_StatusWidth[STATUSDATA2] = iWidthData2;
		m_StatusWidth[STATUSINFOTEXT] = -1;
		RECT rcStatus = { 0 };
		::GetClientRect(m_hWnd, &rcStatus);
		rcStatus.top = rcStatus.bottom - GetStatusHeight();

		// Tell the window it needs to paint
		::RedrawWindow(m_hWnd, &rcStatus, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
	}

	void CBaseDialog::OnSize(UINT/* nType*/, int cx, int cy)
	{
		auto hRes = S_OK;
		HDWP hdwp = nullptr;

		WC_D(hdwp, BeginDeferWindowPos(1));

		if (hdwp)
		{
			const auto iHeight = GetStatusHeight();
			const auto iNewCY = cy - iHeight;
			RECT rcStatus = { 0 };
			::GetClientRect(m_hWnd, &rcStatus);
			if (rcStatus.bottom - rcStatus.top > iHeight)
			{
				rcStatus.top = rcStatus.bottom - iHeight;
			}
			// Tell the status bar it needs repainting
			::InvalidateRect(m_hWnd, &rcStatus, false);

			if (m_lpFakeSplitter && m_lpFakeSplitter->m_hWnd)
			{
				DeferWindowPos(hdwp, m_lpFakeSplitter->m_hWnd, nullptr, 0, 0, cx, iNewCY, SWP_NOZORDER);
			}

			WC_B(EndDeferWindowPos(hdwp));
		}
	}

	void CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos, _In_ const std::wstring& szMsg)
	{
		if (nPos < STATUSBARNUMPANES) m_StatusMessages[nPos] = szMsg;

		SetStatusWidths();
	}

	void __cdecl CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg)
	{
		UpdateStatusBarText(nPos, uidMsg, strings::emptystring, strings::emptystring, strings::emptystring);
	}

	void __cdecl CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg, ULONG ulParam1)
	{
		auto szParam1 = std::to_wstring(ulParam1);
		UpdateStatusBarText(nPos, uidMsg, szParam1, strings::emptystring, strings::emptystring);
	}

	void __cdecl CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg, std::wstring& szParam1, std::wstring& szParam2, std::wstring& szParam3)
	{
		std::wstring szStatBarString;

		// MAPI Load paths take special handling
		if (uidMsg >= ID_LOADMAPIMENUMIN && uidMsg <= ID_LOADMAPIMENUMAX)
		{
			auto hRes = S_OK;
			MENUITEMINFOW mii = { 0 };
			mii.cbSize = sizeof(MENUITEMINFO);
			mii.fMask = MIIM_DATA;

			WC_B(GetMenuItemInfoW(
				::GetMenu(m_hWnd),
				uidMsg,
				false,
				&mii));
			if (mii.dwItemData)
			{
				const auto lme = reinterpret_cast<ui::LPMENUENTRY>(mii.dwItemData);
				szStatBarString = strings::formatmessage(IDS_LOADMAPISTATUS, lme->m_pName.c_str());
			}
		}
		else
		{
			const auto lpAddInMenu = addin::GetAddinMenuItem(m_hWnd, uidMsg);
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
		auto szTitle = strings::formatmessage(IDS_TITLEBARMESSAGE, m_szTitle.c_str(), szMsg.c_str());

		// set the title bar
		::SetWindowTextW(m_hWnd, szTitle.c_str());
	}

	void CBaseDialog::UpdateTitleBarText() const
	{
		auto szTitle = strings::formatmessage(IDS_TITLEBARPLAIN, m_szTitle.c_str());

		// set the title bar
		::SetWindowTextW(m_hWnd, szTitle.c_str());
	}

	void CBaseDialog::UpdateStatus(HWND hWndHost, __StatusPaneEnum pane, const std::wstring& status)
	{
		(void) ::SendMessage(hWndHost, WM_MFCMAPI_UPDATESTATUSBAR, pane, reinterpret_cast<LPARAM>(status.c_str()));
	}

	// WM_MFCMAPI_UPDATESTATUSBAR
	_Check_return_ LRESULT CBaseDialog::msgOnUpdateStatusBar(WPARAM wParam, LPARAM lParam)
	{
		const auto iPane = static_cast<__StatusPaneEnum>(wParam);
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

	void CBaseDialog::OnHexEditor()
	{
		new editor::CHexEditor(m_lpParent, m_lpMapiObjects);
	}

	std::wstring GetOutlookVersionString()
	{
		auto hRes = S_OK;

		if (!pfnMsiProvideQualifiedComponent || !pfnMsiGetFileVersion) return strings::emptystring;

		std::wstring szOut;

		for (auto i = oqcOfficeBegin; i < oqcOfficeEnd; i++)
		{
			auto b64 = false;
			auto lpszTempPath = GetOutlookPath(g_pszOutlookQualifiedComponents[i], &b64);

			if (!lpszTempPath.empty())
			{
				const auto lpszTempVer = new WCHAR[MAX_PATH];
				const auto lpszTempLang = new WCHAR[MAX_PATH];
				if (lpszTempVer && lpszTempLang)
				{
					UINT ret = 0;
					DWORD dwValueBuf = MAX_PATH;
					WC_D(ret, pfnMsiGetFileVersion(lpszTempPath.c_str(),
						lpszTempVer,
						&dwValueBuf,
						lpszTempLang,
						&dwValueBuf));
					if (ERROR_SUCCESS == ret)
					{
						szOut = strings::formatmessage(IDS_OUTLOOKVERSIONSTRING, lpszTempPath.c_str(), lpszTempVer, lpszTempLang);
						szOut += strings::formatmessage(b64 ? IDS_TRUE : IDS_FALSE);
						szOut += L"\n"; // STRING_OK
					}

					delete[] lpszTempLang;
					delete[] lpszTempVer;
				}
			}
		}

		return szOut;
	}

	void CBaseDialog::OnOutlookVersion()
	{
		auto hRes = S_OK;

		editor::CEditor MyEID(
			this,
			IDS_OUTLOOKVERSIONTITLE,
			NULL,
			CEDITOR_BUTTON_OK);

		auto szVersionString = GetOutlookVersionString();
		if (szVersionString.empty())
		{
			szVersionString = strings::loadstring(IDS_NOOUTLOOK);
		}

		MyEID.InitPane(0, viewpane::TextPane::CreateMultiLinePane(IDS_OUTLOOKVERSIONPROMPT, szVersionString, true));
		WC_H(MyEID.DisplayDialog());
	}

	void CBaseDialog::OnOpenEntryID(_In_opt_ LPSBinary lpBin)
	{
		auto hRes = S_OK;
		if (!m_lpMapiObjects) return;

		editor::CEditor MyEID(
			this,
			IDS_OPENEID,
			IDS_OPENEIDPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyEID.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_EID, strings::BinToHexString(lpBin, false), false));

		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		MyEID.InitPane(1, viewpane::CheckPane::Create(IDS_USEMDB, lpMDB != nullptr, lpMDB == nullptr));

		auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
		if (lpAB) MyEID.InitPane(2, viewpane::CheckPane::Create(IDS_USEAB, true, lpAB == nullptr));
		else MyEID.InitPane(2, viewpane::CheckPane::Create(IDS_USEAB, false, lpAB == nullptr));

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		MyEID.InitPane(3, viewpane::CheckPane::Create(IDS_SESSION, lpMAPISession != nullptr, lpMAPISession == nullptr));

		MyEID.InitPane(4, viewpane::CheckPane::Create(IDS_PASSMAPIMODIFY, false, false));

		MyEID.InitPane(5, viewpane::CheckPane::Create(IDS_PASSMAPINOCACHE, false, false));

		MyEID.InitPane(6, viewpane::CheckPane::Create(IDS_PASSMAPICACHEONLY, false, false));

		MyEID.InitPane(7, viewpane::CheckPane::Create(IDS_EIDBASE64ENCODED, false, false));

		MyEID.InitPane(8, viewpane::CheckPane::Create(IDS_ATTEMPTIADDRBOOKDETAILSCALL, false, lpAB == nullptr));

		MyEID.InitPane(9, viewpane::CheckPane::Create(IDS_EIDISCONTAB, false, false));

		WC_H(MyEID.DisplayDialog());
		if (S_OK != hRes) return;

		// Get the entry ID as a binary
		LPENTRYID lpEnteredEntryID = nullptr;
		LPENTRYID lpEntryID = nullptr;
		size_t cbBin = NULL;
		EC_H(MyEID.GetEntryID(0, MyEID.GetCheck(7), &cbBin, &lpEnteredEntryID));

		if (MyEID.GetCheck(9) && lpEnteredEntryID)
		{
			(void)mapi::UnwrapContactEntryID(static_cast<ULONG>(cbBin), reinterpret_cast<LPBYTE>(lpEnteredEntryID), reinterpret_cast<ULONG*>(&cbBin), reinterpret_cast<LPBYTE*>(&lpEntryID));
		}
		else
		{
			lpEntryID = lpEnteredEntryID;
		}

		if (MyEID.GetCheck(8) && lpAB) // Do IAddrBook->Details here
		{
			auto ulUIParam = reinterpret_cast<ULONG_PTR>(static_cast<void*>(m_hWnd));

			EC_H_CANCEL(lpAB->Details(
				&ulUIParam,
				nullptr,
				nullptr,
				static_cast<ULONG>(cbBin),
				lpEntryID,
				nullptr,
				nullptr,
				nullptr,
				DIALOG_MODAL)); // API doesn't like unicode
		}
		else
		{
			LPUNKNOWN lpUnk = nullptr;
			ULONG ulObjType = NULL;

			EC_H(mapi::CallOpenEntry(
				MyEID.GetCheck(1) ? lpMDB : nullptr,
				MyEID.GetCheck(2) ? lpAB : nullptr,
				nullptr,
				MyEID.GetCheck(3) ? lpMAPISession : nullptr,
				static_cast<ULONG>(cbBin),
				lpEntryID,
				nullptr,
				(MyEID.GetCheck(4) ? MAPI_MODIFY : MAPI_BEST_ACCESS) |
				(MyEID.GetCheck(5) ? MAPI_NO_CACHE : 0) |
				(MyEID.GetCheck(6) ? MAPI_CACHE_ONLY : 0),
				&ulObjType,
				&lpUnk));

			if (lpUnk)
			{
				auto szFlags = smartview::InterpretNumberAsStringProp(ulObjType, PR_OBJECT_TYPE);
				output::DebugPrint(DBGGeneric, L"OnOpenEntryID: Got object (%p) of type 0x%08X = %ws\n", lpUnk, ulObjType, szFlags.c_str());

				LPMAPIPROP lpTemp = nullptr;
				WC_MAPI(lpUnk->QueryInterface(IID_IMAPIProp, reinterpret_cast<LPVOID*>(&lpTemp)));
				if (lpTemp)
				{
					WC_H(DisplayObject(
						lpTemp,
						ulObjType,
						otHierarchy,
						this));
					lpTemp->Release();
				}
				lpUnk->Release();
			}
		}

		delete[] lpEnteredEntryID;
	}

	void CBaseDialog::OnCompareEntryIDs()
	{
		auto hRes = S_OK;
		if (!m_lpMapiObjects) return;

		auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release

		editor::CEditor MyEIDs(
			this,
			IDS_COMPAREEIDS,
			IDS_COMPAREEIDSPROMPTS,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyEIDs.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_EID1, false));
		MyEIDs.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_EID2, false));

		UINT uidDropDown[] = {
		IDS_DDMESSAGESTORE,
		IDS_DDSESSION,
		IDS_DDADDRESSBOOK
		};
		MyEIDs.InitPane(2, viewpane::DropDownPane::Create(IDS_OBJECTFORCOMPAREEID, _countof(uidDropDown), uidDropDown, true));

		MyEIDs.InitPane(3, viewpane::CheckPane::Create(IDS_EIDBASE64ENCODED, false, false));

		WC_H(MyEIDs.DisplayDialog());
		if (S_OK != hRes) return;

		if (0 == MyEIDs.GetDropDown(2) && !lpMDB ||
			1 == MyEIDs.GetDropDown(2) && !lpMAPISession ||
			2 == MyEIDs.GetDropDown(2) && !lpAB)
		{
			error::ErrDialog(__FILE__, __LINE__, IDS_EDCOMPAREEID);
			return;
		}
		// Get the entry IDs as a binary
		LPENTRYID lpEntryID1 = nullptr;
		size_t cbBin1 = NULL;
		EC_H(MyEIDs.GetEntryID(0, MyEIDs.GetCheck(3), &cbBin1, &lpEntryID1));

		LPENTRYID lpEntryID2 = nullptr;
		size_t cbBin2 = NULL;
		EC_H(MyEIDs.GetEntryID(1, MyEIDs.GetCheck(3), &cbBin2, &lpEntryID2));

		ULONG ulResult = NULL;
		switch (MyEIDs.GetDropDown(2))
		{
		case 0: // Message Store
			EC_MAPI(lpMDB->CompareEntryIDs(static_cast<ULONG>(cbBin1), lpEntryID1, static_cast<ULONG>(cbBin2), lpEntryID2, NULL, &ulResult));
			break;
		case 1: // Session
			EC_MAPI(lpMAPISession->CompareEntryIDs(static_cast<ULONG>(cbBin1), lpEntryID1, static_cast<ULONG>(cbBin2), lpEntryID2, NULL, &ulResult));
			break;
		case 2: // Address Book
			EC_MAPI(lpAB->CompareEntryIDs(static_cast<ULONG>(cbBin1), lpEntryID1, static_cast<ULONG>(cbBin2), lpEntryID2, NULL, &ulResult));
			break;
		}

		if (SUCCEEDED(hRes))
		{
			auto szResult = strings::loadstring(ulResult ? IDS_TRUE : IDS_FALSE);
			const auto szRet = strings::formatmessage(IDS_COMPAREEIDBOOL, ulResult, szResult.c_str());

			editor::CEditor Result(
				this,
				IDS_COMPAREEIDSRESULT,
				NULL,
				CEDITOR_BUTTON_OK);
			Result.SetPromptPostFix(szRet);
			(void)Result.DisplayDialog();
		}

		delete[] lpEntryID2;
		delete[] lpEntryID1;
	}

	void CBaseDialog::OnComputeStoreHash()
	{
		auto hRes = S_OK;

		editor::CEditor MyStoreEID(
			this,
			IDS_COMPUTESTOREHASH,
			IDS_COMPUTESTOREHASHPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyStoreEID.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_STOREEID, false));
		MyStoreEID.InitPane(1, viewpane::CheckPane::Create(IDS_EIDBASE64ENCODED, false, false));
		MyStoreEID.InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_FILENAME, false));
		MyStoreEID.InitPane(3, viewpane::CheckPane::Create(IDS_PUBLICFOLDERSTORE, false, false));

		WC_H(MyStoreEID.DisplayDialog());
		if (S_OK != hRes) return;

		// Get the entry ID as a binary
		LPENTRYID lpEntryID = nullptr;
		size_t cbBin = NULL;
		EC_H(MyStoreEID.GetEntryID(0, MyStoreEID.GetCheck(1), &cbBin, &lpEntryID));

		const auto dwHash = mapi::ComputeStoreHash(static_cast<ULONG>(cbBin), reinterpret_cast<LPBYTE>(lpEntryID), nullptr, MyStoreEID.GetStringW(2).c_str(), MyStoreEID.GetCheck(3));
		const auto szHash = strings::formatmessage(IDS_STOREHASHVAL, dwHash);

		editor::CEditor Result(
			this,
			IDS_STOREHASH,
			NULL,
			CEDITOR_BUTTON_OK);
		Result.SetPromptPostFix(szHash);
		(void)Result.DisplayDialog();

		delete[] lpEntryID;
	}

	void CBaseDialog::OnNotificationsOn()
	{
		auto hRes = S_OK;

		if (m_lpBaseAdviseSink || !m_lpMapiObjects) return;

		auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release

		editor::CEditor MyData(
			this,
			IDS_NOTIFICATIONS,
			IDS_NOTIFICATIONSPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(interpretprop::AllFlagsToString(flagNotifEventType, true));
		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_EID, false));
		MyData.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_ULEVENTMASK, false));
		MyData.SetHex(1, fnevNewMail);
		UINT uidDropDown[] = {
		IDS_DDMESSAGESTORE,
		IDS_DDSESSION,
		IDS_DDADDRESSBOOK
		};
		MyData.InitPane(2, viewpane::DropDownPane::Create(IDS_OBJECTFORADVISE, _countof(uidDropDown), uidDropDown, true));

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			if (0 == MyData.GetDropDown(2) && !lpMDB ||
				1 == MyData.GetDropDown(2) && !lpMAPISession ||
				2 == MyData.GetDropDown(2) && !lpAB)
			{
				error::ErrDialog(__FILE__, __LINE__, IDS_EDADVISE);
				return;
			}

			LPENTRYID lpEntryID = nullptr;
			size_t cbBin = NULL;
			WC_H(MyData.GetEntryID(0, false, &cbBin, &lpEntryID));
			// don't actually care if the returning lpEntryID is NULL - Advise can work with that

			m_lpBaseAdviseSink = new mapi::mapiui::CAdviseSink(m_hWnd, nullptr);

			if (m_lpBaseAdviseSink)
			{
				switch (MyData.GetDropDown(2))
				{
				case 0:
					EC_MAPI(lpMDB->Advise(
						static_cast<ULONG>(cbBin),
						lpEntryID,
						MyData.GetHex(1),
						static_cast<IMAPIAdviseSink *>(m_lpBaseAdviseSink),
						&m_ulBaseAdviseConnection));
					m_lpBaseAdviseSink->SetAdviseTarget(lpMDB);
					m_ulBaseAdviseObjectType = MAPI_STORE;
					break;
				case 1:
					EC_MAPI(lpMAPISession->Advise(
						static_cast<ULONG>(cbBin),
						lpEntryID,
						MyData.GetHex(1),
						static_cast<IMAPIAdviseSink *>(m_lpBaseAdviseSink),
						&m_ulBaseAdviseConnection));
					m_ulBaseAdviseObjectType = MAPI_SESSION;
					break;
				case 2:
					EC_MAPI(lpAB->Advise(
						static_cast<ULONG>(cbBin),
						lpEntryID,
						MyData.GetHex(1),
						static_cast<IMAPIAdviseSink *>(m_lpBaseAdviseSink),
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
			delete[] lpEntryID;
		}
	}

	void CBaseDialog::OnNotificationsOff()
	{
		auto hRes = S_OK;

		if (m_ulBaseAdviseConnection && m_lpMapiObjects)
		{
			switch (m_ulBaseAdviseObjectType)
			{
			case MAPI_SESSION:
			{
				auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
				if (lpMAPISession) EC_MAPI(lpMAPISession->Unadvise(m_ulBaseAdviseConnection));
				break;
			}
			case MAPI_STORE:
			{
				auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				if (lpMDB) EC_MAPI(lpMDB->Unadvise(m_ulBaseAdviseConnection));
				break;
			}
			case MAPI_ADDRBOOK:
			{
				auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
				if (lpAB) EC_MAPI(lpAB->Unadvise(m_ulBaseAdviseConnection));
				break;
			}
			}
		}

		if (m_lpBaseAdviseSink) m_lpBaseAdviseSink->Release();
		m_lpBaseAdviseSink = nullptr;
		m_ulBaseAdviseObjectType = NULL;
		m_ulBaseAdviseConnection = NULL;
	}

	void CBaseDialog::OnDispatchNotifications()
	{
		auto hRes = S_OK;

		EC_MAPI(HrDispatchNotifications(NULL));
	}

	_Check_return_ bool CBaseDialog::HandleAddInMenu(WORD wMenuSelect)
	{
		output::DebugPrintEx(DBGAddInPlumbing, CLASS, L"HandleAddInMenu", L"wMenuSelect = 0x%08X\n", wMenuSelect);
		return false;
	}

	_Check_return_ ui::CParentWnd* CBaseDialog::GetParentWnd() const
	{
		return m_lpParent;
	}

	_Check_return_ cache::CMapiObjects* CBaseDialog::GetMapiObjects() const
	{
		return m_lpMapiObjects;
	}
}