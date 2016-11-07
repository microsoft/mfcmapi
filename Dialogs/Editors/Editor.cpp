#include "stdafx.h"
#include <Dialogs/Editors/Editor.h>
#include "UIFunctions.h"
#include "String.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "MyWinApp.h"
#include <Dialogs/AboutDlg.h>

extern CMyWinApp theApp;

static wstring CLASS = L"CEditor";

#define NOLIST 0XFFFFFFFF

#define MAX_WIDTH 1000

// After we compute dialog size minimums, when actually set an initial size, we won't go smaller than this
#define MIN_WIDTH 600
#define MIN_HEIGHT 350

#define INVALIDRANGE(iVal) ((iVal) >= m_lpControls.size())

#define LINES_SCROLL 12

// Use this constuctor for generic data editing
CEditor::CEditor(
	_In_opt_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	ULONG ulButtonFlags) :CMyDialog(IDD_BLANK_DIALOG, pParentWnd)
{
	Constructor(pParentWnd,
		uidTitle,
		uidPrompt,
		ulButtonFlags,
		IDS_ACTION1,
		IDS_ACTION2,
		IDS_ACTION3);
}

CEditor::CEditor(
	_In_opt_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	ULONG ulButtonFlags,
	UINT uidActionButtonText1,
	UINT uidActionButtonText2,
	UINT uidActionButtonText3) :CMyDialog(IDD_BLANK_DIALOG, pParentWnd)
{
	Constructor(pParentWnd,
		uidTitle,
		uidPrompt,
		ulButtonFlags,
		uidActionButtonText1,
		uidActionButtonText2,
		uidActionButtonText3);
}

void CEditor::Constructor(
	_In_opt_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	ULONG ulButtonFlags,
	UINT uidActionButtonText1,
	UINT uidActionButtonText2,
	UINT uidActionButtonText3)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_ulListNum = NOLIST;

	m_bEnableScroll = false;
	m_hWndVertScroll = nullptr;
	m_iScrollClient = 0;
	m_bScrollVisible = false;

	m_bButtonFlags = ulButtonFlags;

	m_iMargin = GetSystemMetrics(SM_CXHSCROLL) / 2 + 1;
	m_iSideMargin = m_iMargin / 2;

	m_iSmallHeightMargin = m_iSideMargin;
	m_iLargeHeightMargin = m_iMargin;
	m_iButtonWidth = 50;
	m_iEditHeight = 0;
	m_iTextHeight = 0;
	m_iButtonHeight = 0;
	m_iMinWidth = 0;
	m_iMinHeight = 0;

	m_uidTitle = uidTitle;
	m_uidPrompt = uidPrompt;
	m_bHasPrompt = m_uidPrompt != NULL;

	m_uidActionButtonText1 = uidActionButtonText1;
	m_uidActionButtonText2 = uidActionButtonText2;
	m_uidActionButtonText3 = uidActionButtonText3;

	m_cButtons = 0;
	if (m_bButtonFlags & CEDITOR_BUTTON_OK) m_cButtons++;
	if (m_bButtonFlags & CEDITOR_BUTTON_ACTION1) m_cButtons++;
	if (m_bButtonFlags & CEDITOR_BUTTON_ACTION2) m_cButtons++;
	if (m_bButtonFlags & CEDITOR_BUTTON_ACTION3) m_cButtons++;
	if (m_bButtonFlags & CEDITOR_BUTTON_CANCEL) m_cButtons++;

	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	auto hRes = S_OK;
	WC_D(m_hIcon, AfxGetApp()->LoadIcon(IDR_MAINFRAME));

	m_pParentWnd = pParentWnd;
	if (!m_pParentWnd)
	{
		m_pParentWnd = GetActiveWindow();
	}
	if (!m_pParentWnd)
	{
		m_pParentWnd = theApp.m_pMainWnd;
	}
	if (!m_pParentWnd)
	{
		DebugPrint(DBGGeneric, L"Editor created with a NULL parent!\n");
	}
}

CEditor::~CEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	DeleteControls();
}

BEGIN_MESSAGE_MAP(CEditor, CMyDialog)
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
	ON_WM_NCHITTEST()
END_MESSAGE_MAP()

LRESULT CEditor::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_HELP:
		DisplayAboutDlg(this);
		return true;
		// I can handle notify messages for my child list control since I am the parent window
		// This makes it easy for me to customize the child control to do what I want
	case WM_NOTIFY:
	{
		auto pHdr = reinterpret_cast<LPNMHDR>(lParam);

		switch (pHdr->code)
		{
		case NM_DBLCLK:
		case NM_RETURN:
			auto pane = dynamic_cast<ListPane*>(GetPane(m_ulListNum));
			if (pane)
			{
				(void)pane->HandleChange(IDD_LISTEDIT);
			}

			return NULL;
		}
		break;
	}
	case WM_COMMAND:
	{
		auto nCode = HIWORD(wParam);
		auto idFrom = LOWORD(wParam);
		if (EN_CHANGE == nCode ||
			CBN_SELCHANGE == nCode ||
			CBN_EDITCHANGE == nCode)
		{
			(void)HandleChange(idFrom);
		}
		else if (BN_CLICKED == nCode)
		{
			switch (idFrom)
			{
			case IDD_EDITACTION1: OnEditAction1(); return NULL;
			case IDD_EDITACTION2: OnEditAction2(); return NULL;
			case IDD_EDITACTION3: OnEditAction3(); return NULL;
			case IDD_RECALCLAYOUT: OnRecalcLayout(); return NULL;
			default: (void)HandleChange(idFrom); break;
			}
		}
		break;
	}
	case WM_ERASEBKGND:
	{
		RECT rect = { 0 };
		::GetClientRect(m_hWnd, &rect);
		auto hOld = ::SelectObject(reinterpret_cast<HDC>(wParam), GetSysBrush(cBackground));
		auto bRet = ::PatBlt(reinterpret_cast<HDC>(wParam), 0, 0, rect.right - rect.left, rect.bottom - rect.top, PATCOPY);
		::SelectObject(reinterpret_cast<HDC>(wParam), hOld);
		return bRet;
	}
	case WM_MOUSEWHEEL:
	{
		if (!m_bScrollVisible) break;

		static auto s_DeltaTotal = 0;
		auto zDelta = GET_WHEEL_DELTA_WPARAM(wParam);
		s_DeltaTotal += zDelta;

		int nLines = s_DeltaTotal / WHEEL_DELTA;
		s_DeltaTotal -= nLines * WHEEL_DELTA;
		for (auto i = 0; i != abs(nLines); ++i)
		{
			::SendMessage(m_hWnd,
				WM_VSCROLL,
				nLines < 0 ? static_cast<WPARAM>(SB_LINEDOWN) : static_cast<WPARAM>(SB_LINEUP),
				reinterpret_cast<LPARAM>(m_hWndVertScroll));
		}
		break;
	}
	case WM_VSCROLL:
	{
		auto wScrollType = LOWORD(wParam);
		auto hWndScroll = reinterpret_cast<HWND>(lParam);
		SCROLLINFO si = { 0 };

		si.cbSize = sizeof si;
		si.fMask = SIF_ALL;
		::GetScrollInfo(hWndScroll, SB_CTL, &si);

		// Save the position for comparison later on.
		auto yPos = si.nPos;
		switch (wScrollType)
		{
		case SB_TOP:
			si.nPos = si.nMin;
			break;
		case SB_BOTTOM:
			si.nPos = si.nMax;
			break;
		case SB_LINEUP:
			si.nPos -= m_iEditHeight;
			break;
		case SB_LINEDOWN:
			si.nPos += m_iEditHeight;
			break;
		case SB_PAGEUP:
			si.nPos -= si.nPage;
			break;
		case SB_PAGEDOWN:
			si.nPos += si.nPage;
			break;
		case SB_THUMBTRACK:
			si.nPos = si.nTrackPos;
			break;
		default:
			break;
		}

		// Set the position and then retrieve it. Due to adjustments
		// by Windows it may not be the same as the value set.
		si.fMask = SIF_POS;
		::SetScrollInfo(hWndScroll, SB_CTL, &si, TRUE);
		::GetScrollInfo(hWndScroll, SB_CTL, &si);

		// If the position has changed, scroll window and update it.
		if (si.nPos != yPos)
		{
			::ScrollWindowEx(m_ScrollWindow.m_hWnd, 0, yPos - si.nPos, nullptr, nullptr, nullptr, nullptr, SW_SCROLLCHILDREN | SW_INVALIDATE);
		}

		return 0;
	}
	}
	return CMyDialog::WindowProc(message, wParam, lParam);
}

// AddIn functions
void CEditor::SetAddInTitle(wstring szTitle)
{
	m_szAddInTitle = szTitle;
}

void CEditor::SetAddInLabel(ULONG i, wstring szLabel) const
{
	if (INVALIDRANGE(i)) return;
	auto pane = GetPane(i);
	if (pane) pane->SetAddInLabel(szLabel);
}

LRESULT CALLBACK DrawScrollProc(
	_In_ HWND hWnd,
	UINT uMsg,
	WPARAM wParam,
	LPARAM lParam,
	UINT_PTR uIdSubclass,
	DWORD_PTR /*dwRefData*/)
{
	LRESULT lRes = 0;
	if (HandleControlUI(uMsg, wParam, lParam, &lRes)) return lRes;

	switch (uMsg)
	{
	case WM_NCDESTROY:
		RemoveWindowSubclass(hWnd, DrawScrollProc, uIdSubclass);
		return DefSubclassProc(hWnd, uMsg, wParam, lParam);
	}
	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

// The order these controls are created dictates our tab order - be careful moving things around!
BOOL CEditor::OnInitDialog()
{
	auto hRes = S_OK;
	wstring szPrefix;
	auto szPostfix = loadstring(m_uidTitle);
	wstring szFullString;

	auto bRet = CMyDialog::OnInitDialog();

	m_szTitle = szPostfix + m_szAddInTitle;
	::SetWindowTextW(m_hWnd, m_szTitle.c_str());

	SetIcon(m_hIcon, false); // Set small icon - large icon isn't used

	if (m_bHasPrompt)
	{
		if (m_uidPrompt)
		{
			szPrefix = loadstring(m_uidPrompt);
		}
		else
		{
			// Make sure we clear the prefix out or it might show up in the prompt
			szPrefix = emptystring;
		}

		szFullString = szPrefix + m_szPromptPostFix;

		EC_B(m_Prompt.Create(
			WS_CHILD
			| WS_CLIPSIBLINGS
			| ES_MULTILINE
			| ES_READONLY
			| WS_VISIBLE,
			CRect(0, 0, 0, 0),
			this,
			IDC_PROMPT));
		::SetWindowTextW(m_Prompt.GetSafeHwnd(), szFullString.c_str());

		SubclassLabel(m_Prompt.m_hWnd);
	}

	// setup to get button widths
	auto hdc = ::GetDC(m_hWnd);
	if (!hdc) return false; // fatal error
	auto hfontOld = ::SelectObject(hdc, GetSegoeFont());

	CWnd* pParent = this;
	if (m_bEnableScroll)
	{
		m_ScrollWindow.CreateEx(
			0,
			_T("Static"), // STRING_OK
			nullptr,
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
			CRect(0, 0, 0, 0),
			this,
			NULL);
		m_hWndVertScroll = ::CreateWindowEx(
			0,
			_T("SCROLLBAR"), // STRING_OK
			nullptr,
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | SBS_VERT | SBS_RIGHTALIGN,
			0, 0, 0, 0,
			m_hWnd,
			nullptr,
			nullptr,
			nullptr);
		// Subclass static control so we can ensure we're drawing everything right
		SetWindowSubclass(m_ScrollWindow.m_hWnd, DrawScrollProc, 0, reinterpret_cast<DWORD_PTR>(m_ScrollWindow.m_hWnd));
		pParent = &m_ScrollWindow;
	}

	SetMargins(); // Not all margins have been computed yet, but some have and we can use them during Initialize
	for (ULONG i = 0; i < m_lpControls.size(); i++)
	{
		auto pane = GetPane(i);
		if (pane)
		{
			pane->Initialize(i, pParent, hdc);
		}
	}

	if (m_bButtonFlags & CEDITOR_BUTTON_OK)
	{
		auto szOk = loadstring(IDS_OK);
		EC_B(m_OkButton.Create(
			wstringTotstring(szOk).c_str(),
			WS_TABSTOP
			| WS_CHILD
			| WS_CLIPSIBLINGS
			| WS_VISIBLE,
			CRect(0, 0, 0, 0),
			this,
			IDOK));
	}

	if (m_bButtonFlags & CEDITOR_BUTTON_ACTION1)
	{
		auto szActionButtonText1 = loadstring(m_uidActionButtonText1);
		EC_B(m_ActionButton1.Create(
			wstringTotstring(szActionButtonText1).c_str(),
			WS_TABSTOP
			| WS_CHILD
			| WS_CLIPSIBLINGS
			| WS_VISIBLE,
			CRect(0, 0, 0, 0),
			this,
			IDD_EDITACTION1));

		auto sizeText = GetTextExtentPoint32(hdc, szActionButtonText1);
		m_iButtonWidth = max(m_iButtonWidth, sizeText.cx);
	}

	if (m_bButtonFlags & CEDITOR_BUTTON_ACTION2)
	{
		auto szActionButtonText2 = loadstring(m_uidActionButtonText2);
		EC_B(m_ActionButton2.Create(
			wstringTotstring(szActionButtonText2).c_str(),
			WS_TABSTOP
			| WS_CHILD
			| WS_CLIPSIBLINGS
			| WS_VISIBLE,
			CRect(0, 0, 0, 0),
			this,
			IDD_EDITACTION2));

		auto sizeText = GetTextExtentPoint32(hdc, szActionButtonText2);
		m_iButtonWidth = max(m_iButtonWidth, sizeText.cx);
	}

	if (m_bButtonFlags & CEDITOR_BUTTON_ACTION3)
	{
		auto szActionButtonText3 = loadstring(m_uidActionButtonText3);
		EC_B(m_ActionButton3.Create(
			wstringTotstring(szActionButtonText3).c_str(),
			WS_TABSTOP
			| WS_CHILD
			| WS_CLIPSIBLINGS
			| WS_VISIBLE,
			CRect(0, 0, 0, 0),
			this,
			IDD_EDITACTION3));

		auto sizeText = GetTextExtentPoint32(hdc, szActionButtonText3);
		m_iButtonWidth = max(m_iButtonWidth, sizeText.cx);
	}

	if (m_bButtonFlags & CEDITOR_BUTTON_CANCEL)
	{
		auto szCancel = loadstring(IDS_CANCEL);
		EC_B(m_CancelButton.Create(
			wstringTotstring(szCancel).c_str(),
			WS_TABSTOP
			| WS_CHILD
			| WS_CLIPSIBLINGS
			| WS_VISIBLE,
			CRect(0, 0, 0, 0),
			this,
			IDCANCEL));
	}

	// tear down from our width computations
	(void)::SelectObject(hdc, hfontOld);
	::ReleaseDC(m_hWnd, hdc);

	m_iButtonWidth += m_iMargin;

	// Compute some constants
	m_iEditHeight = GetEditHeight(m_hWnd);
	m_iTextHeight = GetTextHeight(m_hWnd);
	m_iButtonHeight = m_iTextHeight + GetSystemMetrics(SM_CYEDGE) * 2;

	OnSetDefaultSize();

	// Size according to our defaults, respecting minimums
	EC_B(SetWindowPos(
		NULL,
		0,
		0,
		max(MIN_WIDTH, m_iMinWidth),
		max(MIN_HEIGHT, m_iMinHeight),
		SWP_NOZORDER | SWP_NOMOVE));

	return bRet;
}

void CEditor::OnOK()
{
	// save data from the UI back into variables that we can query later
	for (auto pane : m_lpControls)
	{
		if (pane)
		{
			pane->CommitUIValues();
		}
	}

	CMyDialog::OnOK();
}

// This should work whether the editor is active/displayed or not
_Check_return_ GUID CEditor::GetSelectedGUID(ULONG iControl, bool bByteSwapped) const
{
	auto pane = dynamic_cast<DropDownPane*>(GetPane(iControl));
	if (pane)
	{
		return pane->GetSelectedGUID(bByteSwapped);
	}

	return {0};
}

_Check_return_ HRESULT CEditor::DisplayDialog()
{
	auto hRes = S_OK;
	INT_PTR iDlgRet = 0;

	EC_D_DIALOG(DoModal());

	switch (iDlgRet)
	{
	case -1:
		DebugPrint(DBGGeneric, L"Dialog box could not be created!\n");
		MessageBox(_T("Dialog box could not be created!")); // STRING_OK
		return MAPI_E_CALL_FAILED;
	case IDABORT:
	case IDCANCEL:
		return MAPI_E_USER_CANCEL;
	case IDOK:
		return S_OK;
	default:
		return HRESULT_FROM_WIN32(static_cast<unsigned long>(iDlgRet));
	}
}

// Push current margin settings to the view panes
void CEditor::SetMargins() const
{
	for (auto pane : m_lpControls)
	{
		if (pane)
		{
			pane->SetMargins(
				m_iMargin,
				m_iSideMargin,
				m_iTextHeight,
				m_iSmallHeightMargin,
				m_iLargeHeightMargin,
				m_iButtonHeight,
				m_iEditHeight);
		}
	}
}

int ComputePromptWidth(CEdit* lpPrompt, HDC hdc, int iMaxWidth, int* iLineCount)
{
	if (!lpPrompt) return 0;
	auto cx = 0;
	CRect OldRect;
	lpPrompt->GetRect(OldRect);
	// Make the edit rect big so we can get an accurate line count
	lpPrompt->SetRectNP(CRect(0, 0, MAX_WIDTH, iMaxWidth));

	auto iPromptLineCount = lpPrompt->GetLineCount();

	for (auto i = 0; i < iPromptLineCount; i++)
	{
		// length of line i:
		auto len = lpPrompt->LineLength(lpPrompt->LineIndex(i));
		if (len)
		{
			auto szLine = new TCHAR[len + 1];
			memset(szLine, 0, len + 1);

			lpPrompt->GetLine(i, szLine, len);

			int iWidth = LOWORD(::GetTabbedTextExtent(hdc, szLine, len, 0, nullptr));
			delete[] szLine;
			cx = max(cx, iWidth);
		}
	}

	lpPrompt->SetRectNP(OldRect); // restore the old edit rectangle

	iPromptLineCount++; // Add one for an extra line of whitespace
	if (iLineCount) *iLineCount = iPromptLineCount;
	return cx;
}

int ComputeCaptionWidth(HDC hdc, wstring szTitle, int iMargin)
{
	auto sizeTitle = GetTextExtentPoint32(hdc, szTitle);
	auto iCaptionWidth = sizeTitle.cx + iMargin; // Allow for some whitespace between the caption and buttons

	auto iIconWidth = GetSystemMetrics(SM_CXFIXEDFRAME) + GetSystemMetrics(SM_CXSMICON);
	auto iButtonsWidth = GetSystemMetrics(SM_CXBORDER) + 3 * GetSystemMetrics(SM_CYSIZE);
	auto iCaptionMargin = GetSystemMetrics(SM_CXFIXEDFRAME) + GetSystemMetrics(SM_CXBORDER);

	iCaptionWidth += iIconWidth + iButtonsWidth + iCaptionMargin;

	return iCaptionWidth;
}

// Computes good width and height in pixels
// Good is defined as big enough to display all elements at a minimum size, including title
_Check_return_ SIZE CEditor::ComputeWorkArea(SIZE sScreen)
{
	SIZE sArea = { 0 };

	// Figure a good width (cx)
	auto cx = 0;

	auto hdc = ::GetDC(m_hWnd);
	auto hfontOld = ::SelectObject(hdc, GetSegoeFont());

	auto iPromptLineCount = 0;
	if (m_bHasPrompt)
	{
		cx = max(cx, ComputePromptWidth(&m_Prompt, hdc, sScreen.cy, &iPromptLineCount));
	}

	SetMargins();
	// width
	for (auto pane : m_lpControls)
	{
		if (pane)
		{
			cx = max(cx, pane->GetMinWidth(hdc));
		}
	}

	if (m_bEnableScroll)
	{
		cx += GetSystemMetrics(SM_CXVSCROLL) + 2 * GetSystemMetrics(SM_CXFIXEDFRAME);
	}

	(void) ::SelectObject(hdc, hfontOld);

	// Throw all that work out if we have enough buttons
	cx = max(cx, (int)(m_cButtons * m_iButtonWidth + m_iMargin * (m_cButtons - 1)));

	// cx now contains the width of the widest prompt string or control
	// Add a margin around that to frame our controls in the client area:
	cx += 2 * m_iSideMargin;

	// Check that we're wide enough to handle our caption
	cx = max(cx, ComputeCaptionWidth(hdc, m_szTitle, m_iMargin));
	::ReleaseDC(m_hWnd, hdc);
	// Done figuring a good width (cx)

	// Figure a good height (cy)
	auto cy = 2 * m_iMargin; // margins top and bottom
	cy += iPromptLineCount * m_iTextHeight; // prompt text
	cy += m_iButtonHeight; // Button height
	cy += m_iMargin; // add a little height between the buttons and our edit controls

	auto iControlHeight = 0;
	for (auto pane : m_lpControls)
	{
		if (pane)
		{
			iControlHeight +=
				pane->GetFixedHeight() +
				pane->GetLines() * m_iEditHeight;
		}
	}

	if (m_bEnableScroll)
	{
		m_iScrollClient = iControlHeight;
		cy += LINES_SCROLL * m_iEditHeight;
	}
	else
	{
		cy += iControlHeight;
	}
	// Done figuring a good height (cy)

	sArea.cx = cx;
	sArea.cy = cy;
	return sArea;
}

void CEditor::OnSetDefaultSize()
{
	auto hRes = S_OK;

	CRect rcMaxScreen;
	EC_B(SystemParametersInfo(SPI_GETWORKAREA, NULL, static_cast<LPVOID>(static_cast<LPRECT>(rcMaxScreen)), NULL));
	auto cxFullScreen = rcMaxScreen.Width();
	auto cyFullScreen = rcMaxScreen.Height();

	SIZE sScreen = { cxFullScreen, cyFullScreen };
	auto sArea = ComputeWorkArea(sScreen);
	// inflate the rectangle according to the title bar, border, etc...
	CRect MyRect(0, 0, sArea.cx, sArea.cy);
	// Add width and height for the nonclient frame - all previous calculations were done without it
	// This is a call to AdjustWindowRectEx
	CalcWindowRect(MyRect);

	if (MyRect.Width() > cxFullScreen || MyRect.Height() > cyFullScreen)
	{
		// small screen - tighten things up a bit and try again
		m_iMargin = 1;
		m_iSideMargin = 1;
		sArea = ComputeWorkArea(sScreen);
		MyRect.left = 0;
		MyRect.top = 0;
		MyRect.right = sArea.cx;
		MyRect.bottom = sArea.cy;
		CalcWindowRect(MyRect);
	}

	// worst case, don't go bigger than the screen
	m_iMinWidth = MyRect.Width();
	m_iMinHeight = MyRect.Height();
	m_iMinWidth = min(m_iMinWidth, cxFullScreen);
	m_iMinHeight = min(m_iMinHeight, cyFullScreen);

	// worst case v2, don't go bigger than MAX_WIDTH pixels wide
	m_iMinWidth = min(m_iMinWidth, MAX_WIDTH);
}

void CEditor::OnGetMinMaxInfo(_Inout_ MINMAXINFO* lpMMI)
{
	lpMMI->ptMinTrackSize.x = m_iMinWidth;
	lpMMI->ptMinTrackSize.y = m_iMinHeight;
}

// Recalculates our line heights and window defaults, then redraws the window with the new dimensions
void CEditor::OnRecalcLayout()
{
	OnSetDefaultSize();

	RECT rc = { 0 };
	::GetClientRect(m_hWnd, &rc);
	(void) ::PostMessage(m_hWnd,
		WM_SIZE,
		static_cast<WPARAM>(SIZE_RESTORED),
		static_cast<LPARAM>(MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top)));
}

// Artificially expand our gripper region to make it easier to expand dialogs
_Check_return_ LRESULT CEditor::OnNcHitTest(CPoint point)
{
	CRect gripRect;
	GetWindowRect(gripRect);
	gripRect.left = gripRect.right - ::GetSystemMetrics(SM_CXHSCROLL);
	gripRect.top = gripRect.bottom - ::GetSystemMetrics(SM_CYVSCROLL);
	// Test to see if the cursor is within the 'gripper'
	// area, and tell the system that the user is over
	// the lower right-hand corner if it is.
	if (gripRect.PtInRect(point))
	{
		return HTBOTTOMRIGHT;
	}

	return CMyDialog::OnNcHitTest(point);
}

void CEditor::OnSize(UINT nType, int cx, int cy)
{
	auto hRes = S_OK;
	CMyDialog::OnSize(nType, cx, cy);
	auto iCXMargin = m_iSideMargin;

	auto iFullWidth = cx - 2 * iCXMargin;

	auto iPromptLineCount = 0;
	if (m_bHasPrompt)
	{
		iPromptLineCount = m_Prompt.GetLineCount() + 1; // we allow space for the prompt and one line of whitespace
	}

	auto iCYBottom = cy - m_iButtonHeight - m_iMargin; // Top of Buttons
	auto iCYTop = m_iTextHeight * iPromptLineCount + m_iMargin; // Bottom of prompt

	if (m_bHasPrompt)
	{
		// Position prompt at top
		EC_B(m_Prompt.SetWindowPos(
			nullptr, // z-order
			iCXMargin, // new x
			m_iMargin, // new y
			iFullWidth, // Full width
			m_iTextHeight * iPromptLineCount,
			SWP_NOZORDER));
	}

	if (m_cButtons)
	{
		auto iSlotWidth = m_iButtonWidth + m_iMargin;
		auto iOffset = cx - m_iSideMargin + m_iMargin;
		auto iButton = 0;

		// Position buttons at the bottom, on the right
		if (m_bButtonFlags & CEDITOR_BUTTON_OK)
		{
			EC_B(m_OkButton.SetWindowPos(
				nullptr,
				iOffset - iSlotWidth * (m_cButtons - iButton), // new x
				iCYBottom, // new y
				m_iButtonWidth,
				m_iButtonHeight,
				SWP_NOZORDER));
			iButton++;
		}

		if (m_bButtonFlags & CEDITOR_BUTTON_ACTION1)
		{
			EC_B(m_ActionButton1.SetWindowPos(
				nullptr,
				iOffset - iSlotWidth * (m_cButtons - iButton), // new x
				iCYBottom, // new y
				m_iButtonWidth,
				m_iButtonHeight,
				SWP_NOZORDER));
			iButton++;
		}

		if (m_bButtonFlags & CEDITOR_BUTTON_ACTION2)
		{
			EC_B(m_ActionButton2.SetWindowPos(
				nullptr,
				iOffset - iSlotWidth * (m_cButtons - iButton), // new x
				iCYBottom, // new y
				m_iButtonWidth,
				m_iButtonHeight,
				SWP_NOZORDER));
			iButton++;
		}

		if (m_bButtonFlags & CEDITOR_BUTTON_ACTION3)
		{
			EC_B(m_ActionButton3.SetWindowPos(
				nullptr,
				iOffset - iSlotWidth * (m_cButtons - iButton), // new x
				iCYBottom, // new y
				m_iButtonWidth,
				m_iButtonHeight,
				SWP_NOZORDER));
			iButton++;
		}

		if (m_bButtonFlags & CEDITOR_BUTTON_CANCEL)
		{
			EC_B(m_CancelButton.SetWindowPos(
				nullptr,
				iOffset - iSlotWidth * (m_cButtons - iButton), // new x
				iCYBottom, // new y
				m_iButtonWidth,
				m_iButtonHeight,
				SWP_NOZORDER));
		}
	}

	iCYBottom -= m_iMargin; // add a margin above the buttons
	// at this point, iCYTop and iCYBottom reflect our free space, so we can calc multiline height

	// Calculate how much space a 'line' of a variable height control should be
	auto iLineHeight = 0;
	auto iFixedHeight = 0;
	auto iVariableLines = 0;
	for (auto pane : m_lpControls)
	{
		if (pane)
		{
			iFixedHeight += pane->GetFixedHeight();
			iVariableLines += pane->GetLines();
		}
	}
	if (iVariableLines) iLineHeight = (iCYBottom - iCYTop - iFixedHeight) / iVariableLines;

	// There may be some unaccounted slack space after all this. Compute it so we can give it to a control.
	UINT iSlackSpace = iCYBottom - iCYTop - iFixedHeight - iVariableLines * iLineHeight;

	auto iScrollPos = 0;
	if (m_bEnableScroll)
	{
		if (iCYBottom - iCYTop < m_iScrollClient)
		{
			auto iScrollWidth = GetSystemMetrics(SM_CXVSCROLL);
			iFullWidth -= iScrollWidth;
			::SetWindowPos(m_hWndVertScroll, nullptr, iFullWidth + iCXMargin, iCYTop, iScrollWidth, iCYBottom - iCYTop, SWP_NOZORDER);
			SCROLLINFO si = { 0 };
			si.cbSize = sizeof si;
			si.fMask = SIF_POS;
			::GetScrollInfo(m_hWndVertScroll, SB_CTL, &si);
			iScrollPos = si.nPos;

			si.nMin = 0;
			si.nMax = m_iScrollClient;
			si.nPage = iCYBottom - iCYTop;
			si.fMask = SIF_RANGE | SIF_PAGE;
			::SetScrollInfo(m_hWndVertScroll, SB_CTL, &si, FALSE);
			::ShowScrollBar(m_hWndVertScroll, SB_CTL, TRUE);
			m_bScrollVisible = true;
		}
		else
		{
			::ShowScrollBar(m_hWndVertScroll, SB_CTL, FALSE);
			m_bScrollVisible = false;
		}

		::SetWindowPos(m_ScrollWindow.m_hWnd, nullptr, iCXMargin, iCYTop, iFullWidth, iCYBottom - iCYTop, SWP_NOZORDER);
		iCYTop = -iScrollPos; // We get scrolling for free by adjusting our top
	}

	for (auto pane : m_lpControls)
	{
		// Calculate height for multiline edit boxes and lists
		// If we had any slack space, parcel it out Monopoly house style over the controls
		// This ensures a smooth resize experience
		if (pane)
		{
			auto iViewHeight = 0;
			UINT iLines = pane->GetLines();
			if (iLines)
			{
				iViewHeight = iLines * iLineHeight;
				if (iSlackSpace >= iLines)
				{
					iViewHeight += iLines;
					iSlackSpace -= iLines;
				}
				else if (iSlackSpace)
				{
					iViewHeight += iSlackSpace;
					iSlackSpace = 0;
				}
			}

			auto iControlHeight = iViewHeight + pane->GetFixedHeight();
			pane->SetWindowPos(
				iCXMargin, // x
				iCYTop, // y
				iFullWidth, // width
				iControlHeight); // height
			iCYTop += iControlHeight;
		}
	}
}

void CEditor::DeleteControls()
{
	for (auto pane : m_lpControls)
	{
		if (pane) delete[] pane;
	}

	m_lpControls.clear();
}

// TODO: Use iNum as a proper accessor name
void CEditor::InitPane(ULONG iNum, ViewPane* lpPane)
{
	if (lpPane && lpPane->IsType(CTRL_LISTPANE)) m_ulListNum = iNum;
	m_lpControls.push_back(lpPane);
}

// TODO: Use iPane as a proper accessor name
ViewPane* CEditor::GetPane(ULONG iPane) const
{
	if (iPane < 0 || iPane >= m_lpControls.size()) return nullptr;
	return m_lpControls[iPane];
}

void CEditor::SetPromptPostFix(_In_ wstring szMsg)
{
	m_szPromptPostFix = szMsg;
	m_bHasPrompt = true;
}

// Sets string
void CEditor::SetStringA(ULONG i, string szMsg) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{
		pane->SetStringA(szMsg);
	}
}

// Sets string
void CEditor::SetStringW(ULONG i, wstring szMsg) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{
		pane->SetStringW(szMsg);
	}
}

#ifdef CHECKFORMATPARAMS
#undef SetStringf
#endif

// Updates pane using SetStringW
void CEditor::SetStringf(ULONG i, LPCWSTR szMsg, ...) const
{
	if (szMsg[0])
	{
		va_list argList = nullptr;
		va_start(argList, szMsg);
		SetStringW(i, formatV(szMsg, argList));
		va_end(argList);
	}
	else
	{
		SetStringW(i, L"");
	}
}

// Updates pane using SetStringW
void CEditor::LoadString(ULONG i, UINT uidMsg) const
{
	if (uidMsg)
	{
		SetStringW(i, loadstring(uidMsg));
	}
	else
	{
		SetStringW(i, L"");
	}
}

// Updates pane using SetBinary
void CEditor::SetBinary(ULONG i, _In_opt_count_(cb) LPBYTE lpb, size_t cb) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{
		pane->SetBinary(lpb, cb);
	}
}

void CEditor::SetBinary(ULONG i, _In_ vector<BYTE> bin) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{
		pane->SetBinary(bin.data(), bin.size());
	}
}

void CEditor::SetBinary(ULONG i, _In_ SBinary bin) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{
		pane->SetBinary(bin.lpb, bin.cb);
	}
}

// Updates pane using SetStringW
void CEditor::SetSize(ULONG i, size_t cb) const
{
	SetStringf(i, L"0x%08X = %u", static_cast<int>(cb), static_cast<UINT>(cb)); // STRING_OK
}

// Returns a binary buffer which is represented by the hex string
vector<BYTE> CEditor::GetBinaryUseControl(ULONG i) const
{
	return HexStringToBin(GetEditBoxTextW(i));
}

// converts string in a text(edit) control into an entry ID
// Can base64 decode if needed
// entryID is allocated with new, free with delete[]
_Check_return_ HRESULT CEditor::GetEntryID(ULONG i, bool bIsBase64, _Out_ size_t* lpcbBin, _Out_ LPENTRYID* lppEID) const
{
	if (!lpcbBin || !lppEID) return MAPI_E_INVALID_PARAMETER;

	*lpcbBin = NULL;
	*lppEID = nullptr;

	auto hRes = S_OK;
	auto szString = GetStringW(i);

	if (!szString.empty())
	{
		vector<BYTE> bin;
		if (bIsBase64) // entry was BASE64 encoded
		{
			bin = Base64Decode(szString);
		}
		else // Entry was hexized string
		{
			bin = HexStringToBin(szString);
		}

		*lppEID = reinterpret_cast<LPENTRYID>(ByteVectorToLPBYTE(bin));
		*lpcbBin = bin.size();
	}

	return hRes;
}

void CEditor::SetHex(ULONG i, ULONG ulVal) const
{
	SetStringf(i, L"0x%08X", ulVal); // STRING_OK
}

void CEditor::SetDecimal(ULONG i, ULONG ulVal) const
{
	SetStringf(i, L"%u", ulVal); // STRING_OK
}

void CEditor::SetListString(ULONG iControl, ULONG iListRow, ULONG iListCol, wstring szListString) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(iControl));
	if (pane)
	{
		pane->SetListString(iListRow, iListCol, szListString);
	}
}

_Check_return_ SortListData* CEditor::InsertListRow(ULONG iControl, int iRow, wstring szText) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(iControl));
	if (pane)
	{
		return pane->InsertRow(iRow, szText);
	}

	return nullptr;
}

void CEditor::ClearList(ULONG iControl) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(iControl));
	if (pane)
	{
		pane->ClearList();
	}
}

void CEditor::ResizeList(ULONG iControl, bool bSort) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(iControl));
	if (pane)
	{
		pane->ResizeList(bSort);
	}

}

wstring CEditor::GetStringW(ULONG i) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{
		return pane->GetStringW();
	}

	return emptystring;
}

_Check_return_ string CEditor::GetEditBoxTextA(ULONG iControl) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(iControl));
	if (pane)
	{
		return pane->GetEditBoxTextA();
	}

	return "";
}

_Check_return_ wstring CEditor::GetEditBoxTextW(ULONG iControl) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(iControl));
	if (pane)
	{
		return pane->GetEditBoxTextW();
	}

	return emptystring;
}

_Check_return_ ULONG CEditor::GetHex(ULONG i) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{
		return wstringToUlong(pane->GetStringW(), 16);
	}

	return 0;
}

_Check_return_ ULONG CEditor::GetListCount(ULONG iControl) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(iControl));
	if (pane)
	{
		return pane->GetItemCount();
	}

	return 0;
}

_Check_return_ SortListData* CEditor::GetListRowData(ULONG iControl, int iRow) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(iControl));
	if (pane)
	{
		return pane->GetItemData(iRow);
	}

	return nullptr;
}

_Check_return_ bool CEditor::IsDirty(ULONG iControl) const
{
	auto pane = GetPane(iControl);
	return pane ? (vpDirty == (pane->GetFlags() & vpDirty)) : false;
}

_Check_return_ ULONG CEditor::GetHexUseControl(ULONG i) const
{
	return wstringToUlong(GetEditBoxTextW(i), 16);
}

_Check_return_ ULONG CEditor::GetDecimalUseControl(ULONG i) const
{
	return wstringToUlong(GetEditBoxTextW(i), 10);
}

_Check_return_ ULONG CEditor::GetPropTagUseControl(ULONG i) const
{
	auto szTag = GetEditBoxTextW(i);

	// remove any whitespace or nonsense punctuation
	szTag = CleanString(szTag);

	auto ulTag = wstringToUlong(szTag, 16);

	if (ulTag == NULL) // If we didn't convert, try a lookup
	{
		ulTag = PropNameToPropTag(szTag);
	}

	// Figure if this is a full tag or just an ID
	if (ulTag & PROP_TAG_MASK) // Full prop tag
	{
		return ulTag;
	}

	// Just an ID
	return PROP_TAG(PT_UNSPECIFIED, ulTag);
}

_Check_return_ ULONG CEditor::GetPropTag(ULONG i) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{

		auto ulTag = PropNameToPropTag(pane->GetStringW());

		// Figure if this is a full tag or just an ID
		if (ulTag & PROP_TAG_MASK) // Full prop tag
		{
			return ulTag;
		}

		// Just an ID
		return PROP_TAG(PT_UNSPECIFIED, ulTag);
	}

	return 0;
}

_Check_return_ ULONG CEditor::GetDecimal(ULONG i) const
{
	auto pane = dynamic_cast<TextPane*>(GetPane(i));
	if (pane)
	{
		return wstringToUlong(pane->GetStringW(), 10);
	}

	return 0;
}

_Check_return_ bool CEditor::GetCheck(ULONG i) const
{
	auto pane = dynamic_cast<CheckPane*>(GetPane(i));
	if (pane)
	{
		return pane->GetCheck();
	}

	return false;
}

_Check_return_ int CEditor::GetDropDown(ULONG i) const
{
	auto pane = dynamic_cast<DropDownPane*>(GetPane(i));
	if (pane)
	{
		return pane->GetDropDown();
	}

	return CB_ERR;
}

_Check_return_ DWORD_PTR CEditor::GetDropDownValue(ULONG i) const
{
	auto pane = dynamic_cast<DropDownPane*>(GetPane(i));
	if (pane)
	{
		return pane->GetDropDownValue();
	}

	return 0;
}

void CEditor::InsertColumn(ULONG ulListNum, int nCol, UINT uidText) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(ulListNum));
	if (pane)
	{
		pane->InsertColumn(nCol, uidText);
	}
}

void CEditor::InsertColumn(ULONG ulListNum, int nCol, UINT uidText, ULONG ulPropType) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(ulListNum));
	if (pane)
	{
		pane->InsertColumn(nCol, uidText);
		pane->SetColumnType(nCol, ulPropType);
	}
}

_Check_return_ ULONG CEditor::HandleChange(UINT nID)
{
	if (m_lpControls.empty()) return static_cast<ULONG>(-1);
	for (ULONG i = 0; i < m_lpControls.size(); i++)
	{
		auto pane = GetPane(i);
		if (!pane) continue;
		// Either our change came from one of our top level controls/views
		if (pane->MatchID(nID))
		{
			return i;
		}

		// Or the top level control/view has a control in it that can handle it
		// In which case stop looking.
		// We do not return the control number because this is a button event, not an edit change
		if (-1 != pane->HandleChange(nID))
		{
			return static_cast<ULONG>(-1);
		}
	}

	return static_cast<ULONG>(-1);
}

void CEditor::UpdateButtons() const
{
	for (auto pane : m_lpControls)
	{
		if (pane)
		{
			pane->UpdateButtons();
		}
	}
}

_Check_return_ bool CEditor::OnEditListEntry(ULONG ulListNum) const
{
	auto pane = dynamic_cast<ListPane*>(GetPane(ulListNum));
	if (pane)
	{
		return pane->OnEditListEntry();
	}

	return false;
}

void CEditor::OnEditAction1()
{
	// Not Implemented
}

void CEditor::OnEditAction2()
{
	// Not Implemented
}

void CEditor::OnEditAction3()
{
	// Not Implemented
}

// Will be invoked on both edit button and double-click
// return true to indicate the entry was changed, false to indicate it was not
_Check_return_ bool CEditor::DoListEdit(ULONG /*ulListNum*/, int /*iItem*/, _In_ SortListData* /*lpData*/)
{
	// Not Implemented
	return false;
}

void CEditor::EnableScroll()
{
	m_bEnableScroll = true;
}