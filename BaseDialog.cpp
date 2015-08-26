// BaseDialog.cpp : implementation file
//

#include "stdafx.h"
#include "BaseDialog.h"
#include "MainDlg.h"
#include "FakeSplitter.h"
#include "MapiObjects.h"
#include "ParentWnd.h"
#include "SingleMAPIPropListCtrl.h"
#include "Editor.h"
#include "HexEditor.h"
#include "DbgView.h"
#include "MFCUtilityFunctions.h"
#include "UIFunctions.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "AboutDlg.h"
#include "AdviseSink.h"
#include "ExtraPropTags.h"
#include <msi.h>
#include "ImportProcs.h"
#include "SmartView\SmartView.h"
#include "Options.h"
#include "ParseProperty.h"

static wstring CLASS = L"CBaseDialog";

CBaseDialog::CBaseDialog(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects, // Pass NULL to create a new m_lpMapiObjects,
	ULONG ulAddInContext
	) : CMyDialog()
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;
	EC_B(m_szTitle.LoadString(IDS_BASEDIALOG));
	m_bDisplayingMenuText = false;

	m_lpBaseAdviseSink = NULL;
	m_ulBaseAdviseConnection = NULL;
	m_ulBaseAdviseObjectType = NULL;

	m_bIsAB = false;

	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	EC_D(m_hIcon, AfxGetApp()->LoadIcon(IDR_MAINFRAME));

	m_cRef = 1;
	m_lpPropDisplay = NULL;
	m_lpFakeSplitter = NULL;
	// Let the parent know we have a status bar so we can draw our border correctly
	SetStatusHeight(GetSystemMetrics(SM_CXSIZEFRAME) + GetTextHeight(::GetDesktopWindow()));

	m_lpParent = pParentWnd;
	if (m_lpParent) m_lpParent->AddRef();

	m_lpContainer = NULL;
	m_ulAddInContext = ulAddInContext;
	m_ulAddInMenuItems = NULL;

	m_lpMapiObjects = new CMapiObjects(lpMapiObjects);
} // CBaseDialog::CBaseDialog

CBaseDialog::~CBaseDialog()
{
	TRACE_DESTRUCTOR(CLASS);
	HMENU hMenu = ::GetMenu(this->m_hWnd);
	if (hMenu)
	{
		DeleteMenuEntries(hMenu);
		DestroyMenu(hMenu);
	}

	DestroyWindow();
	OnNotificationsOff();
	if (m_lpContainer) m_lpContainer->Release();
	if (m_lpMapiObjects) m_lpMapiObjects->Release();
	if (m_lpParent) m_lpParent->Release();
} // CBaseDialog::~CBaseDialog

STDMETHODIMP_(ULONG) CBaseDialog::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS, lCount);
	DebugPrint(DBGRefCount, L"CBaseDialog::AddRef(\"%ws\")\n", LPCTSTRToWstring(m_szTitle).c_str());
	return lCount;
} // CBaseDialog::AddRef

STDMETHODIMP_(ULONG) CBaseDialog::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS, lCount);
	DebugPrint(DBGRefCount, L"CBaseDialog::Release(\"%ws\")\n", LPCTSTRToWstring(m_szTitle).c_str());
	if (!lCount) delete this;
	return lCount;
} // CBaseDialog::Release

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
		WORD idFrom = LOWORD(wParam);
		// idFrom is the menu item selected
		if (HandleMenu(idFrom)) return S_OK;
		break;
	}
	case WM_PAINT:
		// Paint the status, then let the rest draw itself.
		DrawStatus(
			m_hWnd,
			GetStatusHeight(),
			(LPCTSTR)m_StatusMessages[STATUSDATA1],
			m_StatusWidth[STATUSDATA1],
			(LPCTSTR)m_StatusMessages[STATUSDATA2],
			m_StatusWidth[STATUSDATA2],
			(LPCTSTR)m_StatusMessages[STATUSINFOTEXT]);
		break;
	} // end switch

	return CMyDialog::WindowProc(message, wParam, lParam);
} // CBaseDialog::WindowProc

BOOL CBaseDialog::OnInitDialog()
{
	UpdateTitleBarText(NULL);

	m_StatusWidth[STATUSDATA1] = 0;
	m_StatusWidth[STATUSDATA2] = 0;
	m_StatusWidth[STATUSINFOTEXT] = -1;

	SetIcon(m_hIcon, false); // Set small icon - large icon isn't used

	m_lpFakeSplitter = new CFakeSplitter(this);

	if (m_lpFakeSplitter)
	{
		m_lpPropDisplay = new CSingleMAPIPropListCtrl(m_lpFakeSplitter, this, m_lpMapiObjects, m_bIsAB);

		if (m_lpPropDisplay)
			m_lpFakeSplitter->SetPaneTwo(m_lpPropDisplay);
	}
	return false;
} // CBaseDialog::OnInitDialog

void CBaseDialog::CreateDialogAndMenu(UINT nIDMenuResource, UINT uiClassMenuResource, UINT uidClassMenuTitle)
{
	DebugPrintEx(DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);

	m_lpszTemplateName = MAKEINTRESOURCE(IDD_BLANK_DIALOG);

	DisplayParentedDialog(NULL, NULL);

	HMENU hMenu = NULL;
	if (nIDMenuResource)
	{
		hMenu = ::LoadMenu(NULL, MAKEINTRESOURCE(nIDMenuResource));
	}
	else
	{
		hMenu = ::CreateMenu();
	}

	HMENU hMenuOld = ::GetMenu(m_hWnd);
	if (hMenuOld) ::DestroyMenu(hMenuOld);
	::SetMenu(m_hWnd, hMenu);

	AddMenu(hMenu, IDR_MENU_PROPERTY, IDS_PROPERTYMENU, (UINT)-1);

	AddMenu(hMenu, uiClassMenuResource, uidClassMenuTitle, (UINT)-1);

	m_ulAddInMenuItems = ExtendAddInMenu(hMenu, m_ulAddInContext);

	AddMenu(hMenu, IDR_MENU_TOOLS, IDS_TOOLSMENU, (UINT)-1);

	HMENU hSub = ::GetSubMenu(hMenu, 0);
	::AppendMenu(hSub, MF_SEPARATOR, NULL, NULL);
	WCHAR szExit[16] = { 0 };
	int iRet = NULL;
	iRet = LoadStringW(GetModuleHandle(NULL),
		IDS_EXIT,
		szExit,
		_countof(szExit));
	::AppendMenuW(hSub, MF_ENABLED | MF_STRING, IDCANCEL, szExit);

	// Make sure the menu background is filled in the right color
	MENUINFO mi = { 0 };
	mi.cbSize = sizeof(MENUINFO);
	mi.fMask = MIM_BACKGROUND;
	mi.hbrBack = GetSysBrush(cBackground);
	::SetMenuInfo(hMenu, &mi);

	ConvertMenuOwnerDraw(hMenu, true);

	// We're done - force our new menu on screen
	DrawMenuBar();
} // CBaseDialog::CreateDialogAndMenu

_Check_return_ bool CBaseDialog::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu, L"CBaseDialog::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
	switch (wMenuSelect)
	{
	case ID_HEXEDITOR: OnHexEditor(); return true;
	case ID_DBGVIEW: DisplayDbgView(m_lpParent); return true;
	case ID_COMPAREENTRYIDS: OnCompareEntryIDs(); return true;
	case ID_OPENENTRYID: OnOpenEntryID(NULL); return true;
	case ID_COMPUTESTOREHASH: OnComputeStoreHash(); return true;
	case ID_COPY: HandleCopy(); return true;
	case ID_PASTE: (void)HandlePaste(); return true;
	case ID_OUTLOOKVERSION: OnOutlookVersion(); return true;
	}
	if (HandleAddInMenu(wMenuSelect)) return true;

	if (m_lpPropDisplay) return m_lpPropDisplay->HandleMenu(wMenuSelect);
	return false;
} // CBaseDialog::HandleMenu

void CBaseDialog::OnInitMenu(_In_opt_ CMenu* pMenu)
{
	bool bMAPIInitialized = false;
	if (m_lpMapiObjects)
	{
		bMAPIInitialized = m_lpMapiObjects->bMAPIInitialized();
	}
	if (pMenu)
	{
		if (m_lpPropDisplay) m_lpPropDisplay->InitMenu(pMenu);
		pMenu->EnableMenuItem(ID_NOTIFICATIONSON, DIM(bMAPIInitialized && !m_lpBaseAdviseSink));
		pMenu->CheckMenuItem(ID_NOTIFICATIONSON, CHECK(m_lpBaseAdviseSink));
		pMenu->EnableMenuItem(ID_NOTIFICATIONSOFF, DIM(m_lpBaseAdviseSink));
		pMenu->EnableMenuItem(ID_DISPATCHNOTIFICATIONS, DIM(bMAPIInitialized));
	}
	CMyDialog::OnInitMenu(pMenu);
} // CBaseDialog::OnInitMenu

// Checks flags on add-in menu items to ensure they should be enabled
// Override to support context sensitive scenarios
void CBaseDialog::EnableAddInMenus(_In_ HMENU hMenu, ULONG ulMenu, _In_ LPMENUITEM /*lpAddInMenu*/, UINT uiEnable)
{
	if (hMenu) ::EnableMenuItem(hMenu, ulMenu, uiEnable);
} // CBaseDialog::EnableAddInMenus

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
} // CBaseDialog::OnMenuSelect

_Check_return_ bool CBaseDialog::HandleKeyDown(UINT nChar, bool bShift, bool bCtrl, bool bMenu)
{
	DebugPrintEx(DBGMenu, CLASS, L"HandleKeyDown", L"nChar = 0x%0X, bShift = 0x%X, bCtrl = 0x%X, bMenu = 0x%X\n",
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
			DisplayDbgView(m_lpParent); return true;
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
		OnDeleteSelectedItem(); return true;
		break;
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
		OnEscHit(); return true;
		break;
	case VK_RETURN:
		DebugPrint(DBGMenu, L"CBaseDialog::HandleKeyDown posting ID_DISPLAYSELECTEDITEM\n");
		PostMessage(WM_COMMAND, ID_DISPLAYSELECTEDITEM, NULL);
		return true;
		break;
	}
	return false;
} // CBaseDialog::HandleKeyDown

// prevent dialog from disappearing on Enter
void CBaseDialog::OnOK()
{
	// Now that my controls capture VK_ENTER...this is unneeded...keep it just in case.
} // CBaseDialog::OnOK

void CBaseDialog::OnCancel()
{
	ShowWindow(SW_HIDE);
	if (m_lpPropDisplay) m_lpPropDisplay->Release();
	m_lpPropDisplay = NULL;
	delete m_lpFakeSplitter;
	m_lpFakeSplitter = NULL;
	Release();
} // CBaseDialog::OnCancel

void CBaseDialog::OnEscHit()
{
	DebugPrintEx(DBGGeneric, CLASS, L"OnEscHit", L"Not implemented\n");
} // CBaseDialog::OnEscHit

void CBaseDialog::OnOptions()
{
	ULONG ulNiceNamesBefore = RegKeys[regkeyDO_COLUMN_NAMES].ulCurDWORD;
	ULONG ulSuppressNotFoundBefore = RegKeys[regkeySUPPRESS_NOT_FOUND].ulCurDWORD;
	bool bNeedPropRefresh = DisplayOptionsDlg(this);
	bool bNiceNamesChanged = ulNiceNamesBefore != RegKeys[regkeyDO_COLUMN_NAMES].ulCurDWORD;
	bool bSuppressNotFoundChanged = ulSuppressNotFoundBefore != RegKeys[regkeySUPPRESS_NOT_FOUND].ulCurDWORD;
	HRESULT hRes = S_OK;
	bool bResetColumns = false;

	if (bNiceNamesChanged || bSuppressNotFoundChanged)
	{
		// We check if this worked so we don't refresh the prop list after resetting the top pane
		// But, if we're a tree view, this won't work at all, so we'll still want to reset props if needed
		bResetColumns = false != ::SendMessage(m_hWnd, WM_MFCMAPI_RESETCOLUMNS, 0, 0);
	}

	if (!bResetColumns && bNeedPropRefresh)
	{
		if (m_lpPropDisplay) WC_H(m_lpPropDisplay->RefreshMAPIPropList());
	}
} // CBaseDialog::OnOptions

void CBaseDialog::OnOpenMainWindow()
{
	CMainDlg* pMain = new CMainDlg(m_lpParent, m_lpMapiObjects);
	if (pMain) pMain->OnOpenMessageStoreTable();
} // CBaseDialog::OnOpenMainWindow

void CBaseDialog::HandleCopy()
{
	DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
} // CBaseDialog::HandleCopy

_Check_return_ bool CBaseDialog::HandlePaste()
{
	DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");
	ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();

	if (m_lpPropDisplay && (ulStatus & BUFFER_PROPTAG) && (ulStatus & BUFFER_SOURCEPROPOBJ))
	{
		m_lpPropDisplay->OnPasteProperty();
		return true;
	}

	return false;
} // CBaseDialog::HandlePaste

void CBaseDialog::OnHelp()
{
	DisplayAboutDlg(this);
} // CBaseDialog::OnHelp

void CBaseDialog::OnDeleteSelectedItem()
{
	DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L" Not Implemented\n");
} // CBaseDialog::OnDeleteSelectedItem

void CBaseDialog::OnRefreshView()
{
	DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L" Not Implemented\n");
} // CBaseDialog::OnRefreshView

void CBaseDialog::OnUpdateSingleMAPIPropListCtrl(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ SortListData* lpListData)
{
	HRESULT hRes = S_OK;
	DebugPrintEx(DBGGeneric, CLASS, L"OnUpdateSingleMAPIPropListCtrl", L"Setting item %p\n", lpMAPIProp);

	if (m_lpPropDisplay)
	{
		WC_H(m_lpPropDisplay->SetDataSource(
			lpMAPIProp,
			lpListData,
			m_bIsAB));
	}
} // CBaseDialog::OnUpdateSingleMAPIPropListCtrl

void CBaseDialog::AddMenu(HMENU hMenuBar, UINT uiResource, UINT uidTitle, UINT uiPos)
{
	HMENU hMenuToAdd = ::LoadMenu(NULL, MAKEINTRESOURCE(uiResource));

	if (hMenuBar && hMenuToAdd)
	{
		CString szTitle;
		(void)szTitle.LoadString(uidTitle);
		::InsertMenu(hMenuBar, uiPos, MF_BYPOSITION | MF_POPUP, (UINT_PTR)hMenuToAdd, (LPCTSTR)szTitle);
		if (IDR_MENU_PROPERTY == uiResource)
		{
			(void)ExtendAddInMenu(hMenuToAdd, MENU_CONTEXT_PROPERTY);
		}
	}
} // CBaseDialog::AddMenu

void CBaseDialog::OnActivate(UINT nState, _In_ CWnd* pWndOther, BOOL bMinimized)
{
	HRESULT hRes = S_OK;
	CMyDialog::OnActivate(nState, pWndOther, bMinimized);
	if (nState == 1 && !bMinimized) EC_B(RedrawWindow());
} // CBaseDialog::OnActivate

void CBaseDialog::SetStatusWidths()
{
	// Get the width of the strings
	int iData1Len = m_StatusMessages[STATUSDATA1].GetLength();
	int iData2Len = m_StatusMessages[STATUSDATA2].GetLength();

	SIZE sizeData1 = { 0 };
	SIZE sizeData2 = { 0 };
	if (iData1Len || iData2Len)
	{
		HDC hdc = ::GetDC(m_hWnd);
		if (hdc)
		{
			HGDIOBJ hfontOld = NULL;
			hfontOld = ::SelectObject(hdc, GetSegoeFontBold());

			if (iData1Len)
			{
				::GetTextExtentPoint32(hdc, (LPCTSTR)m_StatusMessages[STATUSDATA1], iData1Len, &sizeData1);
			}

			if (iData2Len)
			{
				::GetTextExtentPoint32(hdc, (LPCTSTR)m_StatusMessages[STATUSDATA2], iData2Len, &sizeData2);
			}

			hfontOld = ::SelectObject(hdc, hfontOld);
			::ReleaseDC(m_hWnd, hdc);
		}
	}

	int nSpacing = GetSystemMetrics(SM_CXEDGE);

	int iWidthData1 = 0;
	int iWidthData2 = 0;
	if (sizeData1.cx) iWidthData1 = sizeData1.cx + 4 * nSpacing;
	if (sizeData2.cx) iWidthData2 = sizeData2.cx + 4 * nSpacing;

	m_StatusWidth[STATUSDATA1] = iWidthData1;
	m_StatusWidth[STATUSDATA2] = iWidthData2;
	m_StatusWidth[STATUSINFOTEXT] = -1;
	RECT rcStatus = { 0 };
	::GetClientRect(m_hWnd, &rcStatus);
	rcStatus.top = rcStatus.bottom - GetStatusHeight();

	// Tell the window it needs to paint
	::RedrawWindow(m_hWnd, &rcStatus, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
} // CBaseDialog::SetStatusWidths

void CBaseDialog::OnSize(UINT/* nType*/, int cx, int cy)
{
	HRESULT hRes = S_OK;
	HDWP hdwp = NULL;

	WC_D(hdwp, BeginDeferWindowPos(1));

	if (hdwp)
	{
		int iHeight = GetStatusHeight();
		int iNewCY = cy - iHeight;
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
			DeferWindowPos(hdwp, m_lpFakeSplitter->m_hWnd, NULL, 0, 0, cx, iNewCY, SWP_NOZORDER);
		}

		WC_B(EndDeferWindowPos(hdwp));
	}
} // CBaseDialog::OnSize

void CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos, _In_z_ LPCTSTR szMsg)
{
	if (nPos < STATUSBARNUMPANES) m_StatusMessages[nPos] = szMsg;

	SetStatusWidths();
} // CBaseDialog::UpdateStatusBarText

void __cdecl CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg)
{
	UpdateStatusBarText(nPos, uidMsg, NULL, NULL, NULL);
}

void __cdecl CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg, ULONG ulParam1)
{
	CString szParam1;
	szParam1.Format(_T("%u"), ulParam1);
	UpdateStatusBarText(nPos, uidMsg, szParam1, NULL, NULL);
}

void __cdecl CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg, LPCTSTR szParam1, LPCTSTR szParam2, LPCTSTR szParam3)
{
	CString szStatBarString;

	// MAPI Load paths take special handling
	if (uidMsg >= ID_LOADMAPIMENUMIN && uidMsg <= ID_LOADMAPIMENUMAX)
	{
		HRESULT hRes = S_OK;
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
			LPMENUENTRY lme = (LPMENUENTRY)mii.dwItemData;
			szStatBarString.FormatMessage(IDS_LOADMAPISTATUS, lme->m_pName);
		}
	}
	else
	{
		LPMENUITEM lpAddInMenu = GetAddinMenuItem(m_hWnd, uidMsg);
		if (lpAddInMenu && lpAddInMenu->szHelp)
		{
			szStatBarString.Format(_T("%ws"), lpAddInMenu->szHelp); // STRING_OK
		}
		else
		{
			HRESULT hRes = S_OK;
			CString szMsg;
			WC_B(szMsg.LoadString(uidMsg));
			if (FAILED(hRes)) DebugPrintEx(DBGMenu, CLASS, L"UpdateStatusBarText", L"Cannot find menu item 0x%08X\n", uidMsg);

			szStatBarString.FormatMessage(szMsg, szParam1, szParam2, szParam3);
		}
	}

	UpdateStatusBarText(nPos, szStatBarString);
} // CBaseDialog::UpdateStatusBarText

void CBaseDialog::UpdateTitleBarText(_In_opt_z_ LPCTSTR szMsg)
{
	CString szTitle;

	if (szMsg)
	{
		szTitle.FormatMessage(IDS_TITLEBARMESSAGE, (LPCTSTR)m_szTitle, szMsg);
	}
	else
	{
		szTitle.FormatMessage(IDS_TITLEBARPLAIN, (LPCTSTR)m_szTitle);
	}
	// set the title bar
	SetWindowText(szTitle);
} // CBaseDialog::UpdateTitleBarText

// WM_MFCMAPI_UPDATESTATUSBAR
_Check_return_ LRESULT	CBaseDialog::msgOnUpdateStatusBar(WPARAM wParam, LPARAM lParam)
{
	__StatusPaneEnum	iPane = (__StatusPaneEnum)wParam;
	LPCTSTR				szStr = (LPCTSTR)lParam;
	UpdateStatusBarText(iPane, szStr);

	return S_OK;
} // CBaseDialog::msgOnUpdateStatusBar

// WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST
_Check_return_ LRESULT	CBaseDialog::msgOnClearSingleMAPIPropList(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	OnUpdateSingleMAPIPropListCtrl(NULL, NULL);

	return S_OK;
} // CBaseDialog::msgOnClearSingleMAPIPropList

void CBaseDialog::OnHexEditor()
{
	new CHexEditor(m_lpParent, m_lpMapiObjects);
} // CBaseDialog::OnHexEditor

CString GetOutlookVersionString()
{
	HRESULT hRes = S_OK;

	if (!pfnMsiProvideQualifiedComponent || !pfnMsiGetFileVersion) return _T("");

	int i = 0;
	CString szOut;

	for (i = oqcOfficeBegin; i < oqcOfficeEnd; i++)
	{
		bool b64 = false;
		LPTSTR lpszTempPath = GetOutlookPath(g_pszOutlookQualifiedComponents[i], &b64);

		if (lpszTempPath)
		{
			LPTSTR lpszTempVer = NULL;
			LPTSTR lpszTempLang = NULL;
			lpszTempVer = new TCHAR[MAX_PATH];
			lpszTempLang = new TCHAR[MAX_PATH];
			if (lpszTempVer && lpszTempLang)
			{
				UINT ret = 0;
				DWORD dwValueBuf = MAX_PATH;
				WC_D(ret, pfnMsiGetFileVersion(lpszTempPath,
					lpszTempVer,
					&dwValueBuf,
					lpszTempLang,
					&dwValueBuf));
				if (ERROR_SUCCESS == ret)
				{
					szOut.AppendFormat(IDS_OUTLOOKVERSIONSTRING, lpszTempPath, lpszTempVer, lpszTempLang);
					szOut.AppendFormat(b64 ? IDS_TRUE : IDS_FALSE);
					szOut.Append(_T("\n")); // STRING_OK
				}
				delete[] lpszTempLang;
				delete[] lpszTempVer;
			}

			delete[] lpszTempPath;
		}
	}

	return szOut;
} // GetOutlookVersionString

void CBaseDialog::OnOutlookVersion()
{
	HRESULT hRes = S_OK;

	CEditor MyEID(
		this,
		IDS_OUTLOOKVERSIONTITLE,
		NULL,
		1,
		CEDITOR_BUTTON_OK);

	CString szVersionString = GetOutlookVersionString();
	if (szVersionString.IsEmpty())
	{
		WC_B(szVersionString.LoadString(IDS_NOOUTLOOK));
	}
	MyEID.InitPane(0, CreateMultiLinePane(IDS_OUTLOOKVERSIONPROMPT, (LPCTSTR)szVersionString, true));


	WC_H(MyEID.DisplayDialog());
} // CBaseDialog::OnOutlookVersion

void CBaseDialog::OnOpenEntryID(_In_opt_ LPSBinary lpBin)
{
	HRESULT hRes = S_OK;
	if (!m_lpMapiObjects) return;

	CEditor MyEID(
		this,
		IDS_OPENEID,
		IDS_OPENEIDPROMPT,
		10,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyEID.InitPane(0, CreateSingleLinePaneW(IDS_EID, BinToHexString(lpBin, false).c_str(), false));

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	MyEID.InitPane(1, CreateCheckPane(IDS_USEMDB, lpMDB ? true : false, lpMDB ? false : true));

	LPADRBOOK lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
	MyEID.InitPane(2, CreateCheckPane(IDS_USEAB, lpAB ? true : false, lpAB ? false : true));

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	MyEID.InitPane(3, CreateCheckPane(IDS_SESSION, lpMAPISession ? true : false, lpMAPISession ? false : true));

	MyEID.InitPane(4, CreateCheckPane(IDS_PASSMAPIMODIFY, false, false));

	MyEID.InitPane(5, CreateCheckPane(IDS_PASSMAPINOCACHE, false, false));

	MyEID.InitPane(6, CreateCheckPane(IDS_PASSMAPICACHEONLY, false, false));

	MyEID.InitPane(7, CreateCheckPane(IDS_EIDBASE64ENCODED, false, false));

	MyEID.InitPane(8, CreateCheckPane(IDS_ATTEMPTIADDRBOOKDETAILSCALL, false, lpAB ? false : true));

	MyEID.InitPane(9, CreateCheckPane(IDS_EIDISCONTAB, false, false));

	WC_H(MyEID.DisplayDialog());
	if (S_OK != hRes) return;

	// Get the entry ID as a binary
	LPENTRYID lpEnteredEntryID = NULL;
	LPENTRYID lpEntryID = NULL;
	size_t cbBin = NULL;
	EC_H(MyEID.GetEntryID(0, MyEID.GetCheck(7), &cbBin, &lpEnteredEntryID));

	if (MyEID.GetCheck(9) && lpEnteredEntryID)
	{
		(void)UnwrapContactEntryID((ULONG)cbBin, (LPBYTE)lpEnteredEntryID, (ULONG*)&cbBin, (LPBYTE*)&lpEntryID);
	}
	else
	{
		lpEntryID = lpEnteredEntryID;
	}

	if (MyEID.GetCheck(8) && lpAB) // Do IAddrBook->Details here
	{
		ULONG_PTR ulUIParam = (ULONG_PTR)(void*)m_hWnd;

		EC_H_CANCEL(lpAB->Details(
			&ulUIParam,
			NULL,
			NULL,
			(ULONG)cbBin,
			lpEntryID,
			NULL,
			NULL,
			NULL,
			DIALOG_MODAL)); // API doesn't like unicode
	}
	else
	{
		LPUNKNOWN lpUnk = NULL;
		ULONG ulObjType = NULL;

		EC_H(CallOpenEntry(
			MyEID.GetCheck(1) ? lpMDB : 0,
			MyEID.GetCheck(2) ? lpAB : 0,
			NULL,
			MyEID.GetCheck(3) ? lpMAPISession : 0,
			(ULONG)cbBin,
			lpEntryID,
			NULL,
			(MyEID.GetCheck(4) ? MAPI_MODIFY : MAPI_BEST_ACCESS) |
			(MyEID.GetCheck(5) ? MAPI_NO_CACHE : 0) |
			(MyEID.GetCheck(6) ? MAPI_CACHE_ONLY : 0),
			&ulObjType,
			&lpUnk));

		if (lpUnk)
		{
			wstring szFlags = InterpretNumberAsStringProp(ulObjType, PR_OBJECT_TYPE);
			DebugPrint(DBGGeneric, L"OnOpenEntryID: Got object (%p) of type 0x%08X = %ws\n", lpUnk, ulObjType, szFlags.c_str());

			LPMAPIPROP lpTemp = NULL;
			WC_MAPI(lpUnk->QueryInterface(IID_IMAPIProp, (LPVOID*)&lpTemp));
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
} // CBaseDialog::OnOpenEntryID

void CBaseDialog::OnCompareEntryIDs()
{
	HRESULT			hRes = S_OK;
	if (!m_lpMapiObjects) return;

	LPMDB			lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	LPMAPISESSION	lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	LPADRBOOK		lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release

	CEditor MyEIDs(
		this,
		IDS_COMPAREEIDS,
		IDS_COMPAREEIDSPROMPTS,
		4,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyEIDs.InitPane(0, CreateSingleLinePane(IDS_EID1, NULL, false));
	MyEIDs.InitPane(1, CreateSingleLinePane(IDS_EID2, NULL, false));

	UINT uidDropDown[] = {
		IDS_DDMESSAGESTORE,
		IDS_DDSESSION,
		IDS_DDADDRESSBOOK
	};
	MyEIDs.InitPane(2, CreateDropDownPane(IDS_OBJECTFORCOMPAREEID, _countof(uidDropDown), uidDropDown, true));

	MyEIDs.InitPane(3, CreateCheckPane(IDS_EIDBASE64ENCODED, false, false));

	WC_H(MyEIDs.DisplayDialog());
	if (S_OK != hRes) return;

	if ((0 == MyEIDs.GetDropDown(2) && !lpMDB) ||
		(1 == MyEIDs.GetDropDown(2) && !lpMAPISession) ||
		(2 == MyEIDs.GetDropDown(2) && !lpAB))
	{
		ErrDialog(__FILE__, __LINE__, IDS_EDCOMPAREEID);
		return;
	}
	// Get the entry IDs as a binary
	LPENTRYID lpEntryID1 = NULL;
	size_t cbBin1 = NULL;
	EC_H(MyEIDs.GetEntryID(0, MyEIDs.GetCheck(3), &cbBin1, &lpEntryID1));

	LPENTRYID lpEntryID2 = NULL;
	size_t cbBin2 = NULL;
	EC_H(MyEIDs.GetEntryID(1, MyEIDs.GetCheck(3), &cbBin2, &lpEntryID2));

	ULONG ulResult = NULL;
	switch (MyEIDs.GetDropDown(2))
	{
	case 0: // Message Store
		EC_MAPI(lpMDB->CompareEntryIDs((ULONG)cbBin1, lpEntryID1, (ULONG)cbBin2, lpEntryID2, NULL, &ulResult));
		break;
	case 1: // Session
		EC_MAPI(lpMAPISession->CompareEntryIDs((ULONG)cbBin1, lpEntryID1, (ULONG)cbBin2, lpEntryID2, NULL, &ulResult));
		break;
	case 2: // Address Book
		EC_MAPI(lpAB->CompareEntryIDs((ULONG)cbBin1, lpEntryID1, (ULONG)cbBin2, lpEntryID2, NULL, &ulResult));
		break;
	}

	if (SUCCEEDED(hRes))
	{
		CString szRet;
		CString szResult;
		EC_B(szResult.LoadString(ulResult ? IDS_TRUE : IDS_FALSE));
		szRet.FormatMessage(IDS_COMPAREEIDBOOL, ulResult, szResult);

		CEditor Result(
			this,
			IDS_COMPAREEIDSRESULT,
			NULL,
			(ULONG)0,
			CEDITOR_BUTTON_OK);
		Result.SetPromptPostFix(szRet);
		(void)Result.DisplayDialog();
	}

	delete[] lpEntryID2;
	delete[] lpEntryID1;
} // CBaseDialog::OnCompareEntryIDs

void CBaseDialog::OnComputeStoreHash()
{
	HRESULT hRes = S_OK;

	CEditor MyStoreEID(
		this,
		IDS_COMPUTESTOREHASH,
		IDS_COMPUTESTOREHASHPROMPT,
		4,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyStoreEID.InitPane(0, CreateSingleLinePane(IDS_STOREEID, NULL, false));
	MyStoreEID.InitPane(1, CreateCheckPane(IDS_EIDBASE64ENCODED, false, false));
	MyStoreEID.InitPane(2, CreateSingleLinePane(IDS_FILENAME, NULL, false));
	MyStoreEID.InitPane(3, CreateCheckPane(IDS_PUBLICFOLDERSTORE, false, false));

	WC_H(MyStoreEID.DisplayDialog());
	if (S_OK != hRes) return;

	// Get the entry ID as a binary
	LPENTRYID lpEntryID = NULL;
	size_t cbBin = NULL;
	EC_H(MyStoreEID.GetEntryID(0, MyStoreEID.GetCheck(1), &cbBin, &lpEntryID));

	DWORD dwHash = ComputeStoreHash((ULONG)cbBin, (LPBYTE)lpEntryID, NULL, MyStoreEID.GetStringW(2), MyStoreEID.GetCheck(3));

	CString szHash;
	szHash.FormatMessage(IDS_STOREHASHVAL, dwHash);

	CEditor Result(
		this,
		IDS_STOREHASH,
		NULL,
		(ULONG)0,
		CEDITOR_BUTTON_OK);
	Result.SetPromptPostFix(szHash);
	(void)Result.DisplayDialog();

	delete[] lpEntryID;
} // CBaseDialog::OnComputeStoreHash

void CBaseDialog::OnNotificationsOn()
{
	HRESULT			hRes = S_OK;

	if (m_lpBaseAdviseSink || !m_lpMapiObjects) return;

	LPMDB			lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	LPMAPISESSION	lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	LPADRBOOK		lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release

	CEditor MyData(
		this,
		IDS_NOTIFICATIONS,
		IDS_NOTIFICATIONSPROMPT,
		3,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.SetPromptPostFix(AllFlagsToString(flagNotifEventType, true));
	MyData.InitPane(0, CreateSingleLinePane(IDS_EID, NULL, false));
	MyData.InitPane(1, CreateSingleLinePane(IDS_ULEVENTMASK, NULL, false));
	MyData.SetHex(1, fnevNewMail);
	UINT uidDropDown[] = {
		IDS_DDMESSAGESTORE,
		IDS_DDSESSION,
		IDS_DDADDRESSBOOK
	};
	MyData.InitPane(2, CreateDropDownPane(IDS_OBJECTFORADVISE, _countof(uidDropDown), uidDropDown, true));

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		if ((0 == MyData.GetDropDown(2) && !lpMDB) ||
			(1 == MyData.GetDropDown(2) && !lpMAPISession) ||
			(2 == MyData.GetDropDown(2) && !lpAB))
		{
			ErrDialog(__FILE__, __LINE__, IDS_EDADVISE);
			return;
		}

		LPENTRYID	lpEntryID = NULL;
		size_t		cbBin = NULL;
		WC_H(MyData.GetEntryID(0, false, &cbBin, &lpEntryID));
		// don't actually care if the returning lpEntryID is NULL - Advise can work with that

		m_lpBaseAdviseSink = new CAdviseSink(m_hWnd, NULL);

		if (m_lpBaseAdviseSink)
		{
			switch (MyData.GetDropDown(2))
			{
			case 0:
				EC_MAPI(lpMDB->Advise(
					(ULONG)cbBin,
					lpEntryID,
					MyData.GetHex(1),
					(IMAPIAdviseSink *)m_lpBaseAdviseSink,
					&m_ulBaseAdviseConnection));
				m_lpBaseAdviseSink->SetAdviseTarget(lpMDB);
				m_ulBaseAdviseObjectType = MAPI_STORE;
				break;
			case 1:
				EC_MAPI(lpMAPISession->Advise(
					(ULONG)cbBin,
					lpEntryID,
					MyData.GetHex(1),
					(IMAPIAdviseSink *)m_lpBaseAdviseSink,
					&m_ulBaseAdviseConnection));
				m_ulBaseAdviseObjectType = MAPI_SESSION;
				break;
			case 2:
				EC_MAPI(lpAB->Advise(
					(ULONG)cbBin,
					lpEntryID,
					MyData.GetHex(1),
					(IMAPIAdviseSink *)m_lpBaseAdviseSink,
					&m_ulBaseAdviseConnection));
				m_lpBaseAdviseSink->SetAdviseTarget(lpAB);
				m_ulBaseAdviseObjectType = MAPI_ADDRBOOK;
				break;
			}

			if (SUCCEEDED(hRes))
			{
				if (0 == MyData.GetDropDown(2) && lpMDB)
				{
					// Try to trigger some RPC to get the notifications going
					LPSPropValue lpProp = NULL;
					WC_MAPI(HrGetOneProp(
						lpMDB,
						PR_TEST_LINE_SPEED,
						&lpProp));
					if (MAPI_E_NOT_FOUND == hRes)
					{
						// We're not on an Exchange server. We don't need to generate RPC after all.
						hRes = S_OK;
					}
					MAPIFreeBuffer(lpProp);
				}
			}
			else // if we failed to advise, then we don't need the advise sink object
			{
				if (m_lpBaseAdviseSink) m_lpBaseAdviseSink->Release();
				m_lpBaseAdviseSink = NULL;
				m_ulBaseAdviseObjectType = NULL;
				m_ulBaseAdviseConnection = NULL;
			}
		}
		delete[] lpEntryID;
	}
} // CBaseDialog::OnNotificationsOn

void CBaseDialog::OnNotificationsOff()
{
	HRESULT hRes = S_OK;

	if (m_ulBaseAdviseConnection && m_lpMapiObjects)
	{
		switch (m_ulBaseAdviseObjectType)
		{
		case MAPI_SESSION:
		{
			LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
			if (lpMAPISession) EC_MAPI(lpMAPISession->Unadvise(m_ulBaseAdviseConnection));
			break;
		}
		case MAPI_STORE:
		{
			LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			if (lpMDB) EC_MAPI(lpMDB->Unadvise(m_ulBaseAdviseConnection));
			break;
		}
		case MAPI_ADDRBOOK:
		{
			LPADRBOOK lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
			if (lpAB) EC_MAPI(lpAB->Unadvise(m_ulBaseAdviseConnection));
			break;
		}
		}
	}

	if (m_lpBaseAdviseSink) m_lpBaseAdviseSink->Release();
	m_lpBaseAdviseSink = NULL;
	m_ulBaseAdviseObjectType = NULL;
	m_ulBaseAdviseConnection = NULL;
} // CBaseDialog::OnNotificationsOff

void CBaseDialog::OnDispatchNotifications()
{
	HRESULT hRes = S_OK;

	EC_MAPI(HrDispatchNotifications(NULL));
} // CBaseDialog::OnDispatchNotifications

_Check_return_ bool CBaseDialog::HandleAddInMenu(WORD wMenuSelect)
{
	DebugPrintEx(DBGAddInPlumbing, CLASS, L"HandleAddInMenu", L"wMenuSelect = 0x%08X\n", wMenuSelect);
	return false;
} // CBaseDialog::HandleAddInMenu

_Check_return_ CParentWnd* CBaseDialog::GetParentWnd()
{
	return m_lpParent;
} // CBaseDialog::GetParentWnd

_Check_return_ CMapiObjects* CBaseDialog::GetMapiObjects()
{
	return m_lpMapiObjects;
} // CBaseDialog::GetMapiObjects