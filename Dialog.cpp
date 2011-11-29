// Editor.cpp : implementation file
//

#include "stdafx.h"
#include "Dialog.h"
#include "UIFunctions.h"
#include "ImportProcs.h"

static TCHAR* CLASS = _T("CMyDialog");

CMyDialog::CMyDialog():CDialog()
{
	Constructor();
} // CMyDialog::CMyDialog

CMyDialog::CMyDialog(UINT nIDTemplate, CWnd* pParentWnd):CDialog(nIDTemplate,pParentWnd)
{
	Constructor();
} // CMyDialog::CMyDialog

void CMyDialog::Constructor()
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpNonModalParent = NULL;
	m_hwndCenteringWindow = NULL;
	m_iAutoCenterWidth = NULL;
	m_iStatusHeight = 0;
} // CMyDialog::Constructor

CMyDialog::~CMyDialog()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpNonModalParent) m_lpNonModalParent->Release();
} // CMyDialog::~CMyDialog

BEGIN_MESSAGE_MAP(CMyDialog, CDialog)
	ON_WM_MEASUREITEM()
	ON_WM_DRAWITEM()
END_MESSAGE_MAP()

void CMyDialog::SetStatusHeight(int iHeight)
{
	m_iStatusHeight = iHeight;
} // CMyDialog::SetStatusHeight

int CMyDialog::GetStatusHeight()
{
	return m_iStatusHeight;
} // CMyDialog::GetStatusHeight

#define WM_NCUAHDRAWCAPTION     0x00AE
#define WM_NCUAHDRAWFRAME       0x00AF

_Check_return_ LRESULT CMyDialog::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_NCUAHDRAWCAPTION:
	case WM_NCUAHDRAWFRAME:
		return 0;
		break;
	case WM_NCACTIVATE:
		// Pass -1 to DefWindowProc to signal we do not want our client repainted.
		lParam =-1;
		DrawWindowFrame(m_hWnd, !!wParam, m_iStatusHeight);
		break;
	case WM_SETTEXT:
		CDialog::WindowProc(message, wParam, lParam);
		DrawWindowFrame(m_hWnd, true, m_iStatusHeight);
		return true;
		break;
	case WM_NCPAINT:
		CDialog::WindowProc(message, wParam, lParam);
		DrawWindowFrame(m_hWnd, true, m_iStatusHeight);
		return 0;
		break;
	case WM_ERASEBKGND:
		return true;
	case WM_CREATE:
		if (pfnSetWindowTheme) (void) pfnSetWindowTheme(m_hWnd, L"", L"");
		break;
	case WM_CTLCOLORSTATIC:
	case WM_CTLCOLOREDIT:
		{
			HDC hdc = (HDC) wParam;
			if (hdc)
			{
				::SetTextColor(hdc, MyGetSysColor(cText));
				::SetBkMode(hdc, TRANSPARENT);
				::SelectObject(hdc, GetSegoeFont());
			}
			return (LRESULT) GetSysBrush(cBackground);
		}
		break;
	case WM_NOTIFY:
		{
			LRESULT lResult = NULL;
			LPNMHDR pHdr = (LPNMHDR) lParam;

			switch (pHdr->code)
			{
			// Paint Buttons
			case NM_CUSTOMDRAW:
				if (CustomDrawButton(pHdr, &lResult)) return lResult;
			}
			break;
		}
	} // end switch
	return CDialog::WindowProc(message, wParam, lParam);
} // CMyDialog::WindowProc

void CMyDialog::DisplayParentedDialog(CParentWnd* lpNonModalParent, UINT iAutoCenterWidth)
{
	HRESULT hRes = S_OK;
	m_iAutoCenterWidth = iAutoCenterWidth;
	m_lpszTemplateName = MAKEINTRESOURCE(IDD_BLANK_DIALOG);

	m_lpNonModalParent = lpNonModalParent;
	if (m_lpNonModalParent) m_lpNonModalParent->AddRef();

	m_hwndCenteringWindow = GetActiveWindow();

	HINSTANCE hInst = AfxFindResourceHandle(m_lpszTemplateName, RT_DIALOG);
	HRSRC hResource = NULL;
	EC_D(hResource,::FindResource(hInst, m_lpszTemplateName, RT_DIALOG));
	if (hResource)
	{
		HGLOBAL hTemplate = NULL;
		EC_D(hTemplate,LoadResource(hInst, hResource));
		if (hTemplate)
		{
			LPCDLGTEMPLATE lpDialogTemplate = (LPCDLGTEMPLATE)LockResource(hTemplate);
			EC_B(CreateDlgIndirect(lpDialogTemplate, m_lpNonModalParent, hInst));
		}
	}
} // CMyDialog::DisplayParentedDialog

// MFC will call this function to check if it ought to center the dialog
// We'll tell it no, but also place the dialog where we want it.
_Check_return_ BOOL CMyDialog::CheckAutoCenter()
{
	// Make the editor wider - OnSize will fix the height for us
	if (m_iAutoCenterWidth)
	{
		SetWindowPos(NULL,0,0,m_iAutoCenterWidth,0,NULL);
	}
	CenterWindow(m_hwndCenteringWindow);
	return false;
} // CMyDialog::CheckAutoCenter

// Measure menu item widths
void CMyDialog::OnMeasureItem(int /*nIDCtl*/, _In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct)
{
	MeasureItem(lpMeasureItemStruct);
} // CMyDialog::OnMeasureItem

// Draw menu items
void CMyDialog::OnDrawItem(int /*nIDCtl*/, _In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	DrawItem(lpDrawItemStruct);
} // CMyDialog::OnDrawItem