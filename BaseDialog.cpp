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
#include "MFCUtilityFunctions.h"
#include "MAPIFunctions.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "AboutDlg.h"
#include "AdviseSink.h"
#include "PropTagArray.h"
#include <msi.h>
#include "ImportProcs.h"

static TCHAR* CLASS = _T("CBaseDialog");

CBaseDialog::CBaseDialog(
				 CParentWnd* pParentWnd,
				 CMapiObjects* lpMapiObjects, // Pass NULL to create a new m_lpMapiObjects,
				 ULONG ulAddInContext
				 ) : CDialog()
{
	TRACE_CONSTRUCTOR(CLASS);
	m_szTitle.LoadString(IDS_BASEDIALOG);
	m_bDisplayingMenuText = false;

	m_lpBaseAdviseSink = NULL;
	m_ulBaseAdviseConnection = NULL;
	m_ulBaseAdviseObjectType = NULL;

	m_bIsAB = false;

	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	HRESULT hRes = S_OK;
	EC_D(m_hIcon,AfxGetApp()->LoadIcon(IDR_MAINFRAME));

	m_cRef = 1;
	m_lpPropDisplay = NULL;
	m_lpFakeSplitter = NULL;

	m_lpParent = pParentWnd;
	if (m_lpParent) m_lpParent->AddRef();

	m_lpContainer = NULL;
	m_ulAddInContext = ulAddInContext;
	m_ulAddInMenuItems = NULL;

	m_lpMapiObjects = new CMapiObjects(lpMapiObjects);
}

CBaseDialog::~CBaseDialog()
{
	TRACE_DESTRUCTOR(CLASS);
	DestroyWindow();
	OnNotificationsOff();
	if (m_lpContainer) m_lpContainer->Release();
	if (m_lpMapiObjects) m_lpMapiObjects->Release();
	if (m_lpParent) m_lpParent->Release();
}

STDMETHODIMP_(ULONG) CBaseDialog::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	DebugPrint(DBGRefCount,_T("CBaseDialog::AddRef(\"%s\")\n"),m_szTitle);
	return lCount;
}

STDMETHODIMP_(ULONG) CBaseDialog::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	DebugPrint(DBGRefCount,_T("CBaseDialog::Release(\"%s\")\n"),m_szTitle);
	if (!lCount) delete this;
	return lCount;
}

BEGIN_MESSAGE_MAP(CBaseDialog, CDialog)
	ON_WM_ACTIVATE()
	ON_WM_INITMENU()
	ON_WM_MENUSELECT()
	ON_WM_SIZE()

	ON_COMMAND(ID_OPTIONS, OnOptions)
	ON_COMMAND(ID_OPENMAINWINDOW, OnOpenMainWindow)
	ON_COMMAND(ID_HELP,OnHelp)

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
	} // end switch
	return CDialog::WindowProc(message,wParam,lParam);
}

BOOL CBaseDialog::OnInitDialog()
{
	HRESULT hRes = S_OK;

	UpdateTitleBarText(NULL);

	EC_B(m_StatusBar.Create(
		WS_CHILD
		| WS_CLIPSIBLINGS
		| WS_VISIBLE
		| CCS_BOTTOM
		| SBARS_SIZEGRIP,
		CRect(0,0,0,0),
		this,
		IDC_STATUS_BAR));

	int StatusWidth[STATUSBARNUMPANES] = {0,0,-1};

	EC_B(m_StatusBar.SetParts(STATUSBARNUMPANES,StatusWidth));

	SetIcon(m_hIcon, FALSE); // Set small icon - large icon isn't used

	m_lpFakeSplitter = new CFakeSplitter(this);

	if (m_lpFakeSplitter)
	{
		m_lpPropDisplay = new CSingleMAPIPropListCtrl(m_lpFakeSplitter,this,m_lpMapiObjects,m_bIsAB);

		if (m_lpPropDisplay)
			m_lpFakeSplitter->SetPaneTwo(m_lpPropDisplay);
	}

	return false;
}

void CBaseDialog::CreateDialogAndMenu(UINT nIDMenuResource)
{
	HRESULT hRes = S_OK;
	CMenu MenuToAdd;

	DebugPrintEx(DBGCreateDialog,CLASS,_T("CreateDialogAndMenu"),_T("id = 0x%X\n"),nIDMenuResource);

	m_lpszTemplateName = MAKEINTRESOURCE(IDD_BLANK_DIALOG);

	HINSTANCE hInst = AfxFindResourceHandle(m_lpszTemplateName, RT_DIALOG);
	HRSRC hResource = NULL;
	EC_D(hResource,::FindResource(hInst, m_lpszTemplateName, RT_DIALOG));
	HGLOBAL hTemplate = NULL;
	EC_D(hTemplate,LoadResource(hInst, hResource));
	LPCDLGTEMPLATE lpDialogTemplate = (LPCDLGTEMPLATE)LockResource(hTemplate);
	EC_B(CreateDlgIndirect(lpDialogTemplate, m_lpParent, hInst));

	if (nIDMenuResource)
	{
		EC_B(MenuToAdd.LoadMenu(nIDMenuResource));
	}
	else
	{
		EC_B(MenuToAdd.CreateMenu());
	}

	EC_B(SetMenu(&MenuToAdd));

	// We add a different file menu if the custom menu already has ID_DISPLAYSELECTEDITEM on it
	HMENU hMenu = ::GetMenu(this->m_hWnd);

	BOOL bOpenExists  = FALSE;
	if (hMenu)
	{
		bOpenExists = (GetMenuState(
			hMenu,
			ID_DISPLAYSELECTEDITEM,
			MF_BYCOMMAND) == -1)? false:true;

	}
	if (bOpenExists)
		AddMenu(IDR_MENU_FILE,IDS_FILEMENU,0);
	else
		AddMenu(IDR_MENU_FILE_OPEN,IDS_FILEMENU,0);

	AddMenu(IDR_MENU_PROPERTY,IDS_PROPERTYMENU,(UINT)-1);

	AddMenu(IDR_MENU_OTHER,IDS_OTHERMENU,(UINT)-1);

	m_ulAddInMenuItems = ExtendAddInMenu(hMenu, m_ulAddInContext);

	// Detach the CMenu object from the menu so cleanup of the CMenu won't affect us
	MenuToAdd.Detach();
} // CBaseDialog::CreateDialogAndMenu

BOOL CBaseDialog::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu,_T("CBaseDialog::HandleMenu wMenuSelect = 0x%X = %d\n"),wMenuSelect,wMenuSelect);
	switch (wMenuSelect)
	{
	case ID_HEXEDITOR: OnHexEditor(); return true;
	case ID_COMPAREENTRYIDS: OnCompareEntryIDs(); return true;
	case ID_OPENENTRYID: OnOpenEntryID(NULL); return true;
	case ID_COMPUTESTOREHASH: OnComputeStoreHash(); return true;
	case ID_ENCODEID: OnEncodeID(); return true;
	case ID_COPY: HandleCopy(); return true;
	case ID_PASTE: HandlePaste(); return true;
	case ID_OUTLOOKVERSION: OnOutlookVersion(); return true;
	}
	if (HandleAddInMenu(wMenuSelect)) return true;

	if (m_lpPropDisplay) return m_lpPropDisplay->HandleMenu(wMenuSelect);
	return false;
}

void CBaseDialog::OnInitMenu(CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpPropDisplay) m_lpPropDisplay->InitMenu(pMenu);
		pMenu->EnableMenuItem(ID_NOTIFICATIONSON,DIM(!m_lpBaseAdviseSink));
		pMenu->CheckMenuItem(ID_NOTIFICATIONSON,CHECK(m_lpBaseAdviseSink));
		pMenu->EnableMenuItem(ID_NOTIFICATIONSOFF,DIM(m_lpBaseAdviseSink));
	}
	CDialog::OnInitMenu(pMenu);
}

// Checks flags on add-in menu items to ensure they should be enabled
// Override to support context sensitive scenarios
void CBaseDialog::EnableAddInMenus(CMenu* pMenu, ULONG ulMenu, LPMENUITEM /*lpAddInMenu*/, UINT uiEnable)
{
	if (pMenu) pMenu->EnableMenuItem(ulMenu,uiEnable);
} // CBaseDialog::EnableAddInMenus

// Help strings can be found in mfcmapi.rc2
// Will preserve the existing text in the right status pane, restoring it when we stop displaying menus
void CBaseDialog::OnMenuSelect(UINT nItemID, UINT nFlags, HMENU /*hSysMenu*/)
{
	if (!m_bDisplayingMenuText)
	{
		m_szMenuDisplacedText = m_StatusBar.GetText(STATUSRIGHTPANE,NULL);
	}

	if (nItemID && !(nFlags & (MF_SEPARATOR | MF_POPUP)))
	{
		UpdateStatusBarText(STATUSRIGHTPANE,nItemID); // This will LoadString the menu help text for us
		m_bDisplayingMenuText = true;
	}
	else
	{
		m_bDisplayingMenuText = false;
	}
	if (!m_bDisplayingMenuText)
	{
		UpdateStatusBarText(STATUSRIGHTPANE,m_szMenuDisplacedText);
	}

	return;

}

BOOL CBaseDialog::HandleKeyDown(UINT nChar, BOOL bShift, BOOL bCtrl, BOOL bMenu)
{
	DebugPrintEx(DBGMenu,CLASS,_T("HandleKeyDown"),_T("nChar = 0x%0X, bShift = 0x%X, bCtrl = 0x%X, bMenu = 0x%X\n"),
		nChar,bShift, bCtrl, bMenu);
	if (!bMenu) switch (nChar)
	{
		case 'D':
			if (bCtrl && !bShift)
			{
				SetDebugLevel(DBGAll); return true;
			}
			if (bCtrl && bShift)
			{
				SetDebugLevel(DBGNoDebug); return true;
			}
			break;
		case 'H':
			if (bCtrl)
			{
				DisplayAboutDlg(this); return true;
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
				HandlePaste(); return true;
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
		case VK_RETURN:
			DebugPrint(DBGMenu,_T("CBaseDialog::HandleKeyDown posting ID_DISPLAYSELECTEDITEM\n"));
			PostMessage(WM_COMMAND,ID_DISPLAYSELECTEDITEM,NULL);
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
	m_lpPropDisplay = NULL;
	delete m_lpFakeSplitter;
	m_lpFakeSplitter = NULL;
	Release();
}

void CBaseDialog::OnEscHit()
{
	DebugPrintEx(DBGGeneric,CLASS,_T("OnEscHit"),_T("Not implemented\n"));
}

void CBaseDialog::OnOptions()
{
	HRESULT	hRes = S_OK;
	DebugPrintEx(DBGGeneric,CLASS,_T("OnOptions"),_T("Building option sheet - creating editor\n"));

	CString szProduct;
	CString szPrompt;
	szProduct.LoadString(ID_PRODUCTNAME);
	szPrompt.FormatMessage(IDS_SETOPTSPROMPT,szProduct);

	CEditor MyData(
		this,
		IDS_SETOPTS,
		NULL,
		NumRegOptionKeys,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.SetPromptPostFix(szPrompt);

	DebugPrintEx(DBGGeneric,CLASS,_T("OnOptions"),_T("Building option sheet - adding fields\n"));

	ULONG ulReg = 0;

	for (ulReg = 0 ; ulReg < NumRegOptionKeys ; ulReg++)
	{
		if (regoptCheck == RegKeys[ulReg].ulRegOptType)
		{
			MyData.InitCheck(ulReg,RegKeys[ulReg].uiOptionsPrompt,RegKeys[ulReg].ulCurDWORD,false);
		}
		else if (regoptString == RegKeys[ulReg].ulRegOptType)
		{
			MyData.InitSingleLineSz(ulReg,RegKeys[ulReg].uiOptionsPrompt,RegKeys[ulReg].szCurSTRING,false);
		}
		else if (regoptStringHex == RegKeys[ulReg].ulRegOptType)
		{
			MyData.InitSingleLine(ulReg,RegKeys[ulReg].uiOptionsPrompt,NULL,false);
			MyData.SetHex(ulReg,RegKeys[ulReg].ulCurDWORD);
		}
		else if (regoptStringDec == RegKeys[ulReg].ulRegOptType)
		{
			MyData.InitSingleLine(ulReg,RegKeys[ulReg].uiOptionsPrompt,NULL,false);
			MyData.SetDecimal(ulReg,RegKeys[ulReg].ulCurDWORD);
		}
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("OnOptions"),_T("Done building option sheet - displaying dialog\n"));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		BOOL bNeedPropRefresh = false;
		// need to grab this FIRST
		EC_H(StringCchCopy(RegKeys[regkeyDEBUG_FILE_NAME].szCurSTRING,CCH(RegKeys[regkeyDEBUG_FILE_NAME].szCurSTRING),MyData.GetString(regkeyDEBUG_FILE_NAME)));

		if (MyData.GetHex(regkeyDEBUG_TAG) != RegKeys[regkeyDEBUG_TAG].ulCurDWORD)
		{
			SetDebugLevel(MyData.GetHex(regkeyDEBUG_TAG));
		}

		SetDebugOutputToFile(MyData.GetCheck(regkeyDEBUG_TO_FILE));

		// Remaining options require no special handling - loop through them
		for (ulReg = 0 ; ulReg < NumRegOptionKeys ; ulReg++)
		{
			if (regoptCheck == RegKeys[ulReg].ulRegOptType)
			{
				if (RegKeys[ulReg].bRefresh && RegKeys[ulReg].ulCurDWORD != (ULONG) MyData.GetCheck(ulReg))
				{
					bNeedPropRefresh = true;
				}
				RegKeys[ulReg].ulCurDWORD = MyData.GetCheck(ulReg);
			}
			else if (regoptStringHex == RegKeys[ulReg].ulRegOptType)
			{
				RegKeys[ulReg].ulCurDWORD = MyData.GetHex(ulReg);
			}
			else if (regoptStringDec == RegKeys[ulReg].ulRegOptType)
			{
				RegKeys[ulReg].ulCurDWORD = MyData.GetDecimal(ulReg);
			}
		}

		// Commit our values to the registry
		WriteToRegistry();

		if (bNeedPropRefresh && m_lpPropDisplay)
			WC_H(m_lpPropDisplay->RefreshMAPIPropList());
	}
}

void CBaseDialog::OnOpenMainWindow()
{
	CMainDlg* pMain = new CMainDlg(m_lpParent,m_lpMapiObjects);
	if (pMain) pMain->OnOpenMessageStoreTable();
}

BOOL CBaseDialog::HandleCopy()
{
	DebugPrintEx(DBGGeneric,CLASS,_T("HandleCopy"),_T("\n"));
	return false;
}

BOOL CBaseDialog::HandlePaste()
{
	DebugPrintEx(DBGGeneric,CLASS,_T("HandlePaste"),_T("\n"));
	ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();

	if (m_lpPropDisplay && (ulStatus & BUFFER_PROPTAG) && (ulStatus & BUFFER_SOURCEPROPOBJ))
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
	DebugPrintEx(DBGDeleteSelectedItem,CLASS,_T("OnDeleteSelectedItem"),_T(" Not Implemented\n"));
}

void CBaseDialog::OnRefreshView()
{
	DebugPrintEx(DBGGeneric,CLASS,_T("OnRefreshView"),_T(" Not Implemented\n"));
}

void CBaseDialog::OnUpdateSingleMAPIPropListCtrl(LPMAPIPROP lpMAPIProp, SortListData* lpListData)
{
	HRESULT hRes = S_OK;
	DebugPrintEx(DBGGeneric,CLASS,_T("OnUpdateSingleMAPIPropListCtrl"),_T("Setting item 0x%X\n"),lpMAPIProp);

	if (m_lpPropDisplay)
	{
		WC_H(m_lpPropDisplay->SetDataSource(
			lpMAPIProp,
			lpListData,
			m_bIsAB));
	}
}

void CBaseDialog::AddMenu(UINT uiResource, UINT uidTitle, UINT uiPos)
{
	HRESULT hRes = S_OK;
	CMenu MyMenu;

	EC_B(MyMenu.LoadMenu(uiResource));

	HMENU hMenu = ::GetMenu(this->m_hWnd);

	if (hMenu)
	{
		CString szTitle;
		szTitle.LoadString(uidTitle);
		EC_B(InsertMenu(
			hMenu,
			uiPos,
			MF_BYPOSITION | MF_POPUP,
			(UINT_PTR) MyMenu.m_hMenu,
			szTitle));

		if (IDR_MENU_PROPERTY == uiResource)
		{
			ExtendAddInMenu(MyMenu.m_hMenu, MENU_CONTEXT_PROPERTY);
		}
	}

	DrawMenuBar();

	MyMenu.Detach();
}

void CBaseDialog::OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized)
{
	HRESULT hRes = S_OK;
	CDialog::OnActivate(nState, pWndOther, bMinimized);
	if (nState == 1 && !bMinimized) EC_B(RedrawWindow());
}

void CBaseDialog::SetStatusWidths()
{
	HRESULT hRes = S_OK;

	int StatusWidth[STATUSBARNUMPANES] = {0};

	EC_B(m_StatusBar.GetParts(STATUSBARNUMPANES,StatusWidth));

	int nType = 0;

	// Get the width of the strings
	int iLeftLen = m_StatusBar.GetTextLength(STATUSLEFTPANE, &nType);
	int iMidLen = m_StatusBar.GetTextLength(STATUSMIDDLEPANE, &nType);

	TCHAR* szText = new TCHAR[1 + max(iLeftLen,iMidLen)];

	if (szText)
	{
		CDC* dcSB = m_StatusBar.GetDC();
		CFont* pFont = dcSB->SelectObject(m_StatusBar.GetFont());

		m_StatusBar.GetText(szText, STATUSLEFTPANE, &nType );
		// this call fails miserably if we don't select a font above
		SIZE sizeLeft = dcSB->GetTabbedTextExtent(szText,0,0);

		m_StatusBar.GetText(szText, STATUSMIDDLEPANE, &nType );
		// this call fails miserably if we don't select a font above
		SIZE sizeMid = dcSB->GetTabbedTextExtent(szText,0,0);

		dcSB->SelectObject(pFont);
		m_StatusBar.ReleaseDC(dcSB);

		int nHorz = 0;
		int nVert = 0;
		int nSpacing = 0;

		EC_B(m_StatusBar.GetBorders(nHorz, nVert, nSpacing));

		int iLeftWidth = sizeLeft.cx+4*nSpacing;
		int iMidWidth = sizeMid.cx+4*nSpacing;

		StatusWidth[STATUSLEFTPANE] = iLeftWidth;
		StatusWidth[STATUSMIDDLEPANE] = iLeftWidth + iMidWidth;
		StatusWidth[STATUSRIGHTPANE] = -1;

		EC_B(m_StatusBar.SetParts(STATUSBARNUMPANES,StatusWidth));

		m_StatusBar.Invalidate(); // force a redraw

	}
	delete[] szText;
}

void CBaseDialog::OnSize(UINT/* nType*/, int cx, int cy)
{
	HRESULT hRes = S_OK;

	SetStatusWidths();

	CRect StatusRect;

	// Figure out how tall the status bar is before we move it:
	EC_B(m_StatusBar.GetRect(STATUSRIGHTPANE,&StatusRect));

	int iHeight = StatusRect.Height();
	int iNewCY = cy-iHeight;

	m_StatusBar.MoveWindow(
		0, // new x
		iNewCY, // new y
		cx,
		iHeight,
		FALSE);

	if (m_lpFakeSplitter && m_lpFakeSplitter->m_hWnd)
	{
		m_lpFakeSplitter->MoveWindow(
			0, // new x
			0, // new y
			cx, // new width
			iNewCY-1, // new height
			TRUE);
	}
}

void CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos,LPCTSTR szMsg)
{
	HRESULT hRes = S_OK;
	// set the status bar
	EC_B(m_StatusBar.SetText(szMsg,nPos,0)); // SBT_NOBORDERS??

	SetStatusWidths();
}

void __cdecl CBaseDialog::UpdateStatusBarText(__StatusPaneEnum nPos,UINT uidMsg,...)
{
	CString szStatBarString;

	LPMENUITEM lpAddInMenu = GetAddinMenuItem(m_hWnd,uidMsg);
	if (lpAddInMenu && lpAddInMenu->szHelp)
	{
		szStatBarString.Format(_T("%ws"),lpAddInMenu->szHelp); // STRING_OK
	}
	else
	{
		CString szMsg;
		szMsg.LoadString(uidMsg);

		va_list argList = NULL;
		va_start(argList, uidMsg);
		szStatBarString.FormatV(szMsg, argList);
		va_end(argList);
	}

	UpdateStatusBarText(nPos,szStatBarString);
} // CBaseDialog::UpdateStatusBarText

void CBaseDialog::UpdateTitleBarText(LPCTSTR szMsg)
{
	CString szTitle;

	if (szMsg)
	{
		szTitle.FormatMessage(_T("%1: %2"),(LPCTSTR) m_szTitle,szMsg); // STRING_OK
	}
	else
	{
		szTitle.FormatMessage(_T("%1"),(LPCTSTR) m_szTitle); // STRING_OK
	}
	// set the title bar
	SetWindowText(szTitle);
} // CBaseDialog::UpdateTitleBarText

// WM_MFCMAPI_UPDATESTATUSBAR
LRESULT	CBaseDialog::msgOnUpdateStatusBar(WPARAM wParam, LPARAM lParam)
{
	__StatusPaneEnum	iPane	= (__StatusPaneEnum) wParam;
	LPCTSTR				szStr	= (LPCTSTR) lParam;
	UpdateStatusBarText(iPane, szStr);

	return S_OK;
}

// WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST
LRESULT	CBaseDialog::msgOnClearSingleMAPIPropList(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	OnUpdateSingleMAPIPropListCtrl(NULL,NULL);

	return S_OK;
}

void CBaseDialog::OnHexEditor()
{
	CHexEditor MyHexEditor(this);
	MyHexEditor.DisplayDialog();
} // CBaseDialog::OnHexEditor

// result allocated with new, clean up with delete[]
void GetOutlookVersionString(LPTSTR* lppszPath, LPTSTR* lppszVer, LPTSTR* lppszLang)
{
	HRESULT hRes = S_OK;
	LPTSTR lpszTempPath = NULL;
	LPTSTR lpszTempVer = NULL;
	LPTSTR lpszTempLang = NULL;

	if (lpszTempPath) *lpszTempPath = NULL;
	if (lpszTempVer) *lpszTempPath = NULL;
	if (lpszTempLang) *lpszTempLang = NULL;

	if (!pfnMsiProvideQualifiedComponent || !pfnMsiGetFileVersion) return;

	TCHAR pszaOutlookQualifiedComponents[][MAX_PATH] = {
		_T("{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}"), // O11_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
		_T("{BC174BAD-2F53-4855-A1D5-1D575C19B1EA}"), // O11_CATEGORY_GUID_CORE_OFFICE (debug) // STRING_OK
		_T("{24AAE126-0911-478F-A019-07B875EB9996}")  // O12_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	};
	int nOutlookQualifiedComponents = sizeof(pszaOutlookQualifiedComponents)/sizeof(TCHAR*);
	int i = 0;
	DWORD dwValueBuf = 0;
	UINT ret = 0;

	for (i = 0; i < nOutlookQualifiedComponents; i++)
	{
		WC_D(ret,pfnMsiProvideQualifiedComponent(
			pszaOutlookQualifiedComponents[i],
			_T("outlook.exe"), // STRING_OK
			(DWORD) INSTALLMODE_DEFAULT,
			NULL,
			&dwValueBuf));
		if (ERROR_SUCCESS == ret) break;
	}

	if (ERROR_SUCCESS == ret)
	{
		dwValueBuf += 1;
		lpszTempPath = new TCHAR[dwValueBuf];

		if (lpszTempPath != NULL)
		{
			WC_D(ret,pfnMsiProvideQualifiedComponent(
				pszaOutlookQualifiedComponents[i],
				_T("outlook.exe"), // STRING_OK
				(DWORD) INSTALLMODE_EXISTING,
				lpszTempPath,
				&dwValueBuf));

			if (ERROR_SUCCESS == ret)
			{
				lpszTempVer = new TCHAR[MAX_PATH];
				lpszTempLang = new TCHAR[MAX_PATH];
				dwValueBuf = MAX_PATH;
				if (lpszTempVer && lpszTempLang)
				{
					WC_D(ret,pfnMsiGetFileVersion(lpszTempPath,
						lpszTempVer,
						&dwValueBuf,
						lpszTempLang,
						&dwValueBuf));
					if (ERROR_SUCCESS == ret)
					{
						if (lppszVer)
						{
							*lppszVer = lpszTempVer;
							lpszTempVer = NULL;
						}
					}

					if (lppszPath)
					{
						*lppszPath = lpszTempPath;
						lpszTempPath = NULL;
					}
					if (lppszLang)
					{
						*lppszLang = lpszTempLang;
						lpszTempLang = NULL;
					}
				}
			}
		}
	}

	delete[] lpszTempVer;
	delete[] lpszTempLang;
	delete[] lpszTempPath;
} // GetOutlookVersionString

void CBaseDialog::OnOutlookVersion()
{
	HRESULT hRes = S_OK;
	LPTSTR lpszPath = NULL;
	LPTSTR lpszVer = NULL;
	LPTSTR lpszLang = NULL;
	GetOutlookVersionString(&lpszPath, &lpszVer, &lpszLang);

	CEditor MyEID(
		this,
		IDS_OUTLOOKVERSIONTITLE,
		IDS_OUTLOOKVERSIONPROMPT,
		3,
		CEDITOR_BUTTON_OK);

	MyEID.InitSingleLineSz(0,IDS_OUTLOOKVERSIONLABELPATH,lpszPath,true);
	MyEID.InitSingleLineSz(1,IDS_OUTLOOKVERSIONLABELVER,lpszVer,true);
	MyEID.InitSingleLineSz(2,IDS_OUTLOOKVERSIONLABELLANG,lpszLang,true);
	if (!lpszPath)
	{
		MyEID.LoadString(0,IDS_NOTFOUND);
	}
	if (!lpszVer)
	{
		MyEID.LoadString(1,IDS_NOTFOUND);
	}
	if (!lpszLang)
	{
		MyEID.LoadString(2,IDS_NOTFOUND);
	}
	delete[] lpszPath;
	delete[] lpszVer;

	WC_H(MyEID.DisplayDialog());
} // CBaseDialog::OnOutlookVersion

void CBaseDialog::OnOpenEntryID(LPSBinary lpBin)
{
	HRESULT			hRes = S_OK;
	if (!m_lpMapiObjects) return;

	CEditor MyEID(
		this,
		IDS_OPENEID,
		IDS_OPENEIDPROMPT,
		10,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyEID.InitSingleLineSz(0,IDS_EID,BinToHexString(lpBin,false),false);

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	MyEID.InitCheck(1,IDS_USEMDB,lpMDB?true:false,lpMDB?false:true);

	LPADRBOOK lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
	MyEID.InitCheck(2,IDS_USEAB,lpAB?true:false,lpAB?false:true);

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	MyEID.InitCheck(3,IDS_SESSION,lpMAPISession?true:false,lpMAPISession?false:true);

	MyEID.InitCheck(4,IDS_PASSMAPIMODIFY,false,false);

	MyEID.InitCheck(5,IDS_PASSMAPINOCACHE,false,false);

	MyEID.InitCheck(6,IDS_PASSMAPICACHEONLY,false,false);

	MyEID.InitCheck(7,IDS_EIDBASE64ENCODED,false,false);

	MyEID.InitCheck(8,IDS_ATTEMPTIADDRBOOKDETAILSCALL,false,lpAB?false:true);

	MyEID.InitCheck(9,IDS_EIDISCONTAB,false,false);

	WC_H(MyEID.DisplayDialog());
	if (S_OK != hRes) return;

	// Get the entry ID as a binary
	LPENTRYID lpEnteredEntryID = NULL;
	LPENTRYID lpEntryID = NULL;
	size_t cbBin = NULL;
	EC_H(MyEID.GetEntryID(0,MyEID.GetCheck(7),&cbBin,&lpEnteredEntryID));

	if (MyEID.GetCheck(9) && lpEnteredEntryID)
	{
		LPCONTAB_ENTRYID lpContabEID = (LPCONTAB_ENTRYID)lpEnteredEntryID;
		if (lpContabEID && lpContabEID->cbeid && lpContabEID->abeid)
		{
			cbBin = lpContabEID->cbeid;
			lpEntryID = (LPENTRYID) lpContabEID->abeid;
		}

	}
	else
	{
		lpEntryID = lpEnteredEntryID;
	}

	if (MyEID.GetCheck(8) && lpAB) // Do IAddrBook->Details here
	{
		ULONG_PTR ulUIParam = (ULONG_PTR) (void*) m_hWnd;

		EC_H_CANCEL(lpAB->Details(
			&ulUIParam,
			NULL,
			NULL,
			(ULONG) cbBin,
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
			MyEID.GetCheck(1)?lpMDB:0,
			MyEID.GetCheck(2)?lpAB:0,
			NULL,
			MyEID.GetCheck(3)?lpMAPISession:0,
			(ULONG) cbBin,
			lpEntryID,
			NULL,
			(MyEID.GetCheck(4)?MAPI_MODIFY:MAPI_BEST_ACCESS) |
			(MyEID.GetCheck(5)?MAPI_NO_CACHE:0) |
			(MyEID.GetCheck(6)?MAPI_CACHE_ONLY:0),
			&ulObjType,
			&lpUnk));

		if (lpUnk)
		{
			LPTSTR szFlags = NULL;
			EC_H(InterpretFlags(PROP_ID(PR_OBJECT_TYPE), ulObjType, &szFlags));
			DebugPrint(DBGGeneric,_T("OnOpenEntryID: Got object (0x%08X) of type 0x%08X = %s\n"),lpUnk,ulObjType,szFlags);
			delete[] szFlags;
			szFlags = NULL;

			LPMAPIPROP lpTemp = NULL;
			WC_H(lpUnk->QueryInterface(IID_IMAPIProp,(LPVOID*) &lpTemp));
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

	return;
} // OnOpenEntryID

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
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyEIDs.InitSingleLine(0,IDS_EID1,NULL,false);
	MyEIDs.InitSingleLine(1,IDS_EID2,NULL,false);

	UINT uidDropDown[] = {
		IDS_DDMESSAGESTORE,
			IDS_DDSESSION,
			IDS_DDADDRESSBOOK
	};
	MyEIDs.InitDropDown(2,IDS_OBJECTFORCOMPAREEID,sizeof(uidDropDown)/sizeof(UINT),uidDropDown,true);

	MyEIDs.InitCheck(3,IDS_EIDBASE64ENCODED,false,false);

	WC_H(MyEIDs.DisplayDialog());
	if (S_OK != hRes) return;

	if ((0 == MyEIDs.GetDropDown(2) && !lpMDB) ||
		(1 == MyEIDs.GetDropDown(2) && !lpMAPISession) ||
		(2 == MyEIDs.GetDropDown(2) && !lpAB))
	{
		ErrDialog(__FILE__,__LINE__,IDS_EDCOMPAREEID);
		return;
	}
	// Get the entry IDs as a binary
	LPENTRYID lpEntryID1 = NULL;
	size_t cbBin1 = NULL;
	EC_H(MyEIDs.GetEntryID(0,MyEIDs.GetCheck(3),&cbBin1,&lpEntryID1));

	LPENTRYID lpEntryID2 = NULL;
	size_t cbBin2 = NULL;
	EC_H(MyEIDs.GetEntryID(1,MyEIDs.GetCheck(3),&cbBin2,&lpEntryID2));

	ULONG ulResult = NULL;
	switch (MyEIDs.GetDropDown(2))
	{
	case 0: // Message Store
		EC_H(lpMDB->CompareEntryIDs((ULONG)cbBin1,lpEntryID1,(ULONG)cbBin2,lpEntryID2,NULL,&ulResult));
		break;
	case 1: // Session
		EC_H(lpMAPISession->CompareEntryIDs((ULONG)cbBin1,lpEntryID1,(ULONG)cbBin2,lpEntryID2,NULL,&ulResult));
		break;
	case 2: // Address Book
		EC_H(lpAB->CompareEntryIDs((ULONG)cbBin1,lpEntryID1,(ULONG)cbBin2,lpEntryID2,NULL,&ulResult));
		break;
	}

	if (SUCCEEDED(hRes))
	{
		CString szRet;
		CString szResult;
		szResult.LoadString(ulResult?IDS_TRUE:IDS_FALSE);
		szRet.FormatMessage(_T("0x%1!08X! = %2"),ulResult,szResult); // STRING_OK

		CEditor Result(
			this,
			IDS_COMPAREEIDSRESULT,
			NULL,
			(ULONG) 0,
			CEDITOR_BUTTON_OK);
		Result.SetPromptPostFix(szRet);
		Result.DisplayDialog();
	}

	delete[] lpEntryID2;
	delete[] lpEntryID1;

	return;
} // OnCompareEntryIDs

void CBaseDialog::OnComputeStoreHash()
{
	HRESULT hRes = S_OK;

	CEditor MyStoreEID(
		this,
		IDS_COMPUTESTOREHASH,
		IDS_COMPUTESTOREHASHPROMPT,
		3,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyStoreEID.InitSingleLine(0,IDS_STOREEID,NULL,false);
	MyStoreEID.InitCheck(1,IDS_EIDBASE64ENCODED,false,false);
	MyStoreEID.InitSingleLine(2,IDS_FILENAME,NULL,false);

	WC_H(MyStoreEID.DisplayDialog());
	if (S_OK != hRes) return;

	// Get the entry ID as a binary
	LPENTRYID lpEntryID = NULL;
	size_t cbBin = NULL;
	EC_H(MyStoreEID.GetEntryID(0,MyStoreEID.GetCheck(1),&cbBin,&lpEntryID));

	DWORD dwHash = ComputeStoreHash((ULONG) cbBin,lpEntryID,MyStoreEID.GetStringW(2));

	CString szHash;
	CString szHashStr;
	szHashStr.LoadString(IDS_STOREHASH);
	szHash.FormatMessage(_T("%1 = 0x%2!08X!"),szHashStr,dwHash); // STRING_OK

	CEditor Result(
		this,
		IDS_STOREHASH,
		NULL,
		(ULONG) 0,
		CEDITOR_BUTTON_OK);
	Result.SetPromptPostFix(szHash);
	Result.DisplayDialog();

	delete[] lpEntryID;

	return;
} // CBaseDialog::OnComputeStoreHash

void CBaseDialog::OnEncodeID()
{
	HRESULT hRes = S_OK;

	CEditor MyEID(
		this,
		IDS_ENCODEID,
		IDS_ENCODEIDPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyEID.InitSingleLine(0,IDS_EID,NULL,false);
	MyEID.InitCheck(1,IDS_EIDBASE64ENCODED,false,false);

	WC_H(MyEID.DisplayDialog());
	if (S_OK != hRes) return;

	// Get the entry ID as a binary
	LPENTRYID lpEntryID = NULL;
	size_t cbBin = NULL;
	EC_H(MyEID.GetEntryID(0,MyEID.GetCheck(1),&cbBin,&lpEntryID));

	LPWSTR szEncoded = EncodeID((ULONG) cbBin,lpEntryID);

	if (szEncoded)
	{
		CEditor Result(
			this,
			IDS_ENCODEDID,
			IDS_ENCODEDID,
			(ULONG) 3,
			CEDITOR_BUTTON_OK);
		Result.InitSingleLine(0,IDS_UNISTRING,NULL,true);
		Result.SetStringW(0,szEncoded);
		size_t cchEncoded = NULL;
		WC_H(StringCchLengthW(szEncoded,STRSAFE_MAX_CCH,&cchEncoded));
		Result.InitSingleLine(1,IDS_CCH,NULL,true);
		Result.SetHex(1,(ULONG) cchEncoded);
		Result.InitMultiLine(2,IDS_HEX,NULL,true);
		Result.SetBinary(2,(LPBYTE)szEncoded,cchEncoded);

		Result.DisplayDialog();
		delete[] szEncoded;
	}

	delete[] lpEntryID;

	return;
} // CBaseDialog::OnEncodeID

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
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.SetPromptPostFix(AllFlagsToString(flagNotifEventType,true));
	MyData.InitSingleLine(0,IDS_EID,NULL,false);
	MyData.InitSingleLine(1,IDS_ULEVENTMASK,NULL,false);
	MyData.SetHex(1,fnevNewMail);
	UINT uidDropDown[] = {
		IDS_DDMESSAGESTORE,
			IDS_DDSESSION,
			IDS_DDADDRESSBOOK
	};
	MyData.InitDropDown(2,IDS_OBJECTFORADVISE,sizeof(uidDropDown)/sizeof(UINT),uidDropDown,true);

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		if ((0 == MyData.GetDropDown(2) && !lpMDB) ||
			(1 == MyData.GetDropDown(2) && !lpMAPISession) ||
			(2 == MyData.GetDropDown(2) && !lpAB))
		{
			ErrDialog(__FILE__,__LINE__,IDS_EDADVISE);
			return;
		}

		LPENTRYID	lpEntryID = NULL;
		size_t		cbBin = NULL;
		WC_H(MyData.GetEntryID(0,false,&cbBin,&lpEntryID));
		// don't actually care if the returning lpEntryID is NULL - Advise can work with that

		m_lpBaseAdviseSink = new CAdviseSink(m_hWnd,NULL);

		if (m_lpBaseAdviseSink)
		{
			switch (MyData.GetDropDown(2))
			{
			case 0:
				EC_H(lpMDB->Advise(
					(ULONG) cbBin,
					lpEntryID,
					MyData.GetHex(1),
					(IMAPIAdviseSink *)m_lpBaseAdviseSink,
					&m_ulBaseAdviseConnection));
				m_ulBaseAdviseObjectType = MAPI_STORE;
				break;
			case 1:
				EC_H(lpMAPISession->Advise(
					(ULONG) cbBin,
					lpEntryID,
					MyData.GetHex(1),
					(IMAPIAdviseSink *)m_lpBaseAdviseSink,
					&m_ulBaseAdviseConnection));
				m_ulBaseAdviseObjectType = MAPI_SESSION;
				break;
			case 2:
				EC_H(lpAB->Advise(
					(ULONG) cbBin,
					lpEntryID,
					MyData.GetHex(1),
					(IMAPIAdviseSink *)m_lpBaseAdviseSink,
					&m_ulBaseAdviseConnection));
				m_ulBaseAdviseObjectType = MAPI_ADDRBOOK;
				break;
			}

			if (SUCCEEDED(hRes))
			{
				if (0 == MyData.GetDropDown(2) && lpMDB)
				{
					// Try to trigger some RPC to get the notifications going
					LPSPropValue lpProp = NULL;
					WC_H(HrGetOneProp(
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
				if (lpMAPISession) EC_H(lpMAPISession->Unadvise(m_ulBaseAdviseConnection));
			}
			break;
		case MAPI_STORE:
			{
				LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				if (lpMDB) EC_H(lpMDB->Unadvise(m_ulBaseAdviseConnection));
			}
			break;
		case MAPI_ADDRBOOK:
			{
				LPADRBOOK lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
				if (lpAB) EC_H(lpAB->Unadvise(m_ulBaseAdviseConnection));
			}
			break;
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

	EC_H(HrDispatchNotifications(NULL));
} // CBaseDialog::OnDispatchNotifications

BOOL CBaseDialog::HandleAddInMenu(WORD wMenuSelect)
{
	DebugPrintEx(DBGAddInPlumbing,CLASS,_T("HandleAddInMenu"),_T("wMenuSelect = 0x%08X\n"),wMenuSelect);
	return false;
} // CBaseDialog::HandleAddInMenu

CParentWnd* CBaseDialog::GetParentWnd()
{
	return m_lpParent;
}

CMapiObjects* CBaseDialog::GetMapiObjects()
{
	return m_lpMapiObjects;
}