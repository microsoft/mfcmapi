#include <StdAfx.h>
#include <UI/ViewPane/CheckPane.h>
#include <UI/UIFunctions.h>
#include <core/utility/output.h>

namespace viewpane
{
	std::shared_ptr<CheckPane>
	CheckPane::Create(const int paneID, const UINT uidLabel, const bool bVal, const bool bReadOnly)
	{
		auto pane = std::make_shared<CheckPane>();
		if (pane)
		{
			pane->m_bCheckValue = bVal;
			pane->m_szLabel = strings::loadstring(uidLabel);

			pane->SetReadOnly(bReadOnly);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	int CheckPane::GetMinWidth()
	{
		const auto label = ViewPane::GetMinWidth();
		return max(label, m_iLabelWidth);
	}

	void CheckPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);

		EC_B_S(m_Check.Create(
			nullptr,
			WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | BS_AUTOCHECKBOX | (m_bReadOnly ? WS_DISABLED : 0) |
				BS_NOTIFY,
			CRect(0, 0, 0, 0),
			pParent,
			m_nID));
		m_Check.SetCheck(m_bCheckValue);
		::SetWindowTextW(m_Check.m_hWnd, m_szLabel.c_str());

		const auto sizeText = ui::GetTextExtentPoint32(hdc, m_szLabel);
		const auto check = GetSystemMetrics(SM_CXMENUCHECK);
		const auto edge = check / 5;
		m_iLabelWidth = check + edge + sizeText.cx;

		m_bInitialized = true;
	}

	// CheckPane Layout:
	// Top margin: m_iSmallHeightMargin if requested
	// Header: none
	// CheckPane: m_iButtonHeight
	int CheckPane::GetFixedHeight()
	{
		auto height = m_iButtonHeight;
		if (m_bTopMargin) height += m_iSmallHeightMargin;
		return height;
	}

	HDWP CheckPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int /*width*/,
		_In_ const int height)
	{
		auto curY = y;
		if (m_bTopMargin) curY += m_iSmallHeightMargin;

		hWinPosInfo = ui::DeferWindowPos(
			hWinPosInfo,
			m_Check.GetSafeHwnd(),
			x,
			curY,
			m_iLabelWidth,
			height - (curY - y),
			L"CheckPane::DeferWindowPos");
		return hWinPosInfo;
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

		WC_D_S(::FillRect(hDC, &rcCheck, GetSysBrush(ui::uiColor::Background)));
		if (bFocused)
		{
			ui::FrameRect(hDC, rcCheck, 3, ui::uiColor::Glow);
		}
		else if (bGlow)
		{
			ui::FrameRect(hDC, rcCheck, 2, ui::uiColor::Glow);
		}
		else if (bDisabled)
		{
			ui::FrameRect(hDC, rcCheck, 1, ui::uiColor::FrameUnselected);
		}
		else
		{
			ui::FrameRect(hDC, rcCheck, 1, ui::uiColor::FrameSelected);
		}

		if (bChecked)
		{
			auto rcFill = rcCheck;
			const auto deflate = lEdge;
			InflateRect(&rcFill, -deflate, -deflate);
			FillRect(hDC, &rcFill, GetSysBrush(ui::uiColor::Glow));
		}

		auto rcLabel = rc;
		rcLabel.left = rcCheck.right + lEdge;
		rcLabel.right = rcLabel.left + ui::GetTextExtentPoint32(hDC, szButton).cx;

		output::DebugPrint(
			output::dbgLevel::Draw,
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
			bDisabled ? MyGetSysColor(ui::uiColor::TextDisabled) : MyGetSysColor(ui::uiColor::Text),
			rcLabel,
			false,
			DT_SINGLELINE | DT_VCENTER);
	}

	bool CheckPane::containsWindow(HWND hWnd) const noexcept
	{
		if (m_Check.GetSafeHwnd() == hWnd) return true;
		return m_Header.containsWindow(hWnd);
	}
} // namespace viewpane