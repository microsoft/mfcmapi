#include <StdAfx.h>
#include <UI/ViewPane/CountedTextPane.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>

namespace viewpane
{
	std::shared_ptr<CountedTextPane>
	CountedTextPane::Create(const int paneID, const UINT uidLabel, const bool bReadOnly, const UINT uidCountLabel)
	{
		auto pane = std::make_shared<CountedTextPane>();
		if (pane)
		{
			pane->m_szCountLabel = strings::loadstring(uidCountLabel);
			pane->SetMultiline();
			pane->SetLabel(uidLabel);
			pane->ViewPane::SetReadOnly(bReadOnly);
			pane->m_bCollapsible = true;
			pane->m_paneID = paneID;
		}

		return pane;
	}

	void CountedTextPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		EC_B_S(m_Count.Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE, CRect(0, 0, 0, 0), pParent, IDD_COUNTLABEL));
		::SetWindowTextW(m_Count.m_hWnd, m_szCountLabel.c_str());
		ui::SubclassLabel(m_Count.m_hWnd);
		StyleLabel(m_Count.m_hWnd, ui::uiLabelStyle::PaneHeaderText);

		TextPane::Initialize(pParent, hdc);
	}

	int CountedTextPane::GetMinWidth()
	{
		const auto iLabelWidth = TextPane::GetMinWidth();

		auto szCount = strings::format(
			L"%ws: 0x%08X = %u",
			m_szCountLabel.c_str(),
			static_cast<int>(m_iCount),
			static_cast<UINT>(m_iCount)); // STRING_OK
		::SetWindowTextW(m_Count.m_hWnd, szCount.c_str());

		const auto hdc = GetDC(m_Count.GetSafeHwnd());
		const auto hfontOld = SelectObject(hdc, ui::GetSegoeFont());
		const auto sizeText = ui::GetTextExtentPoint32(hdc, szCount);
		static_cast<void>(SelectObject(hdc, hfontOld));
		ReleaseDC(m_Count.GetSafeHwnd(), hdc);
		m_iCountLabelWidth = sizeText.cx + m_iSideMargin;

		// Button, margin, label, margin, count label
		const auto cx = m_iButtonHeight + m_iSideMargin + iLabelWidth + m_iSideMargin + m_iCountLabelWidth;

		return cx;
	}

	int CountedTextPane::GetFixedHeight()
	{
		auto iHeight = 0;
		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		iHeight += GetLabelHeight();

		if (!m_bCollapsed)
		{
			// Small gap before the edit box
			iHeight += m_iSmallHeightMargin;
		}

		iHeight += m_iSmallHeightMargin; // Bottom margin

		return iHeight;
	}

	int CountedTextPane::GetLines() { return m_bCollapsed ? 0 : LINES_MULTILINEEDIT; }

	HDWP CountedTextPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		auto curY = y;
		const auto labelHeight = GetLabelHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		// Layout our label
		hWinPosInfo = EC_D(HDWP, ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y)));

		if (m_bCollapsed)
		{
			WC_B_S(m_Count.ShowWindow(SW_HIDE));
			WC_B_S(m_EditBox.ShowWindow(SW_HIDE));

			hWinPosInfo = EC_D(
				HDWP, ::DeferWindowPos(hWinPosInfo, m_EditBox.GetSafeHwnd(), nullptr, x, curY, 0, 0, SWP_NOZORDER));
		}
		else
		{
			WC_B_S(m_Count.ShowWindow(SW_SHOW));
			WC_B_S(m_EditBox.ShowWindow(SW_SHOW));

			// Drop the count on top of the label we drew above
			hWinPosInfo = EC_D(
				HDWP,
				::DeferWindowPos(
					hWinPosInfo,
					m_Count.GetSafeHwnd(),
					nullptr,
					x + width - m_iCountLabelWidth,
					curY,
					m_iCountLabelWidth,
					labelHeight,
					SWP_NOZORDER));

			curY += labelHeight + m_iSmallHeightMargin;

			hWinPosInfo = EC_D(
				HDWP,
				::DeferWindowPos(
					hWinPosInfo,
					m_EditBox.GetSafeHwnd(),
					nullptr,
					x,
					curY,
					width,
					height - (curY - y) - m_iSmallHeightMargin,
					SWP_NOZORDER));
		}

		return hWinPosInfo;
	}

	void CountedTextPane::SetCount(size_t iCount) { m_iCount = iCount; }
} // namespace viewpane