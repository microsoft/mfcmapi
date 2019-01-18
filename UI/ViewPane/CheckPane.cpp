#include <StdAfx.h>
#include <UI/ViewPane/CheckPane.h>
#include <UI/UIFunctions.h>
#include <core/utility/output.h>

namespace viewpane
{
	CheckPane* CheckPane::Create(const int paneID, const UINT uidLabel, const bool bVal, const bool bReadOnly)
	{
		auto pane = new (std::nothrow) CheckPane();
		if (pane)
		{
			pane->m_bCheckValue = bVal;
			pane->SetLabel(uidLabel);
			pane->SetReadOnly(bReadOnly);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	int CheckPane::GetMinWidth()
	{
		const auto label = ViewPane::GetMinWidth();
		const auto check = GetSystemMetrics(SM_CXMENUCHECK);
		const auto edge = check / 5;
		output::DebugPrint(
			DBGDraw,
			L"CheckPane::GetMinWidth Label:%d + check:%d + edge:%d = minwidth:%d\n",
			label,
			check,
			edge,
			label + edge + check);
		return label + edge + check;
	}

	int CheckPane::GetFixedHeight() { return m_iButtonHeight; }

	void CheckPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);

		EC_B_S(m_Check.Create(
			NULL,
			WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | BS_AUTOCHECKBOX | (m_bReadOnly ? WS_DISABLED : 0),
			CRect(0, 0, 0, 0),
			pParent,
			m_nID));
		m_Check.SetCheck(m_bCheckValue);
		SetWindowTextW(m_Check.m_hWnd, m_szLabel.c_str());

		m_bInitialized = true;
	}

	void CheckPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		output::DebugPrint(DBGDraw, L"CheckPane::DeferWindowPos x:%d width:%d \n", x, width);
		EC_B_S(::DeferWindowPos(hWinPosInfo, m_Check.GetSafeHwnd(), nullptr, x, y, width, height, SWP_NOZORDER));
	}

	void CheckPane::CommitUIValues()
	{
		m_bCheckValue = 0 != m_Check.GetCheck();
		m_bCommitted = true;
	}

	bool CheckPane::GetCheck() const
	{
		if (!m_bCommitted) return 0 != m_Check.GetCheck();
		return m_bCheckValue;
	}

	void CheckPane::Draw(_In_ HWND hWnd, _In_ HDC hDC, _In_ const RECT& rc, const UINT itemState)
	{
		WCHAR szButton[255];
		GetWindowTextW(hWnd, szButton, _countof(szButton));
		const auto iState = ::SendMessage(hWnd, BM_GETSTATE, NULL, NULL);
		const auto bGlow = (iState & BST_HOT) != 0;
		const auto bChecked = (iState & BST_CHECKED) != 0;
		const auto bDisabled = (itemState & CDIS_DISABLED) != 0;
		const auto bFocused = (itemState & CDIS_FOCUS) != 0;

		const auto lCheck = GetSystemMetrics(SM_CXMENUCHECK);
		const auto lEdge = lCheck / 5;
		auto rcCheck = RECT{0};
		rcCheck.left = rc.left;
		rcCheck.right = rcCheck.left + lCheck;
		rcCheck.top = (rc.bottom - rc.top - lCheck) / 2;
		rcCheck.bottom = rcCheck.top + lCheck;

		FillRect(hDC, &rc, GetSysBrush(ui::cBackground));
		FrameRect(
			hDC,
			&rcCheck,
			GetSysBrush(bDisabled ? ui::cFrameUnselected : bGlow || bFocused ? ui::cGlow : ui::cFrameSelected));
		if (bChecked)
		{
			auto rcFill = rcCheck;
			const auto deflate = lEdge;
			InflateRect(&rcFill, -deflate, -deflate);
			FillRect(hDC, &rcFill, GetSysBrush(ui::cGlow));
		}

		auto rcLabel = rc;
		rcLabel.left = rcCheck.right + lEdge;
		rcLabel.right = rcLabel.left + ui::GetTextExtentPoint32(hDC, szButton).cx;

		output::DebugPrint(
			DBGDraw,
			L"CheckButton::Draw left:%d width:%d checkwidth:%d space:%d labelwidth:%d (scroll:%d 2frame:%d), \"%ws\"\n",
			rc.left,
			rc.right - rc.left,
			rcCheck.right - rcCheck.left,
			lEdge,
			rcLabel.right - rcLabel.left,
			GetSystemMetrics(SM_CXVSCROLL),
			2 * GetSystemMetrics(SM_CXFIXEDFRAME),
			szButton);

		ui::DrawSegoeTextW(
			hDC,
			szButton,
			bDisabled ? MyGetSysColor(ui::cTextDisabled) : MyGetSysColor(ui::cText),
			rcLabel,
			false,
			DT_SINGLELINE | DT_VCENTER);
	}
} // namespace viewpane