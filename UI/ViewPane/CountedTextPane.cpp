#include <StdAfx.h>
#include <UI/ViewPane/CountedTextPane.h>
#include <UI/UIFunctions.h>

namespace viewpane
{
	static std::wstring CLASS = L"CountedTextPane";

	CountedTextPane* CountedTextPane::Create(int paneID, UINT uidLabel, bool bReadOnly, UINT uidCountLabel)
	{
		auto lpPane = new (std::nothrow) CountedTextPane();
		if (lpPane)
		{
			lpPane->m_szCountLabel = strings::loadstring(uidCountLabel);
			lpPane->SetMultiline();
			lpPane->SetLabel(uidLabel, bReadOnly);
			lpPane->m_bCollapsible = true;
			lpPane->m_paneID = paneID;
		}

		return lpPane;
	}

	void CountedTextPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		EC_B_S(m_Count.Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE, CRect(0, 0, 0, 0), pParent, IDD_COUNTLABEL));
		SetWindowTextW(m_Count.m_hWnd, m_szCountLabel.c_str());
		ui::SubclassLabel(m_Count.m_hWnd);
		StyleLabel(m_Count.m_hWnd, ui::lsPaneHeaderText);

		TextPane::Initialize(pParent, hdc);
	}

	int CountedTextPane::GetMinWidth(_In_ HDC hdc)
	{
		const auto iLabelWidth = TextPane::GetMinWidth(hdc);

		auto szCount = strings::format(
			L"%ws: 0x%08X = %u",
			m_szCountLabel.c_str(),
			static_cast<int>(m_iCount),
			static_cast<UINT>(m_iCount)); // STRING_OK
		SetWindowTextW(m_Count.m_hWnd, szCount.c_str());

		const auto sizeText = ui::GetTextExtentPoint32(hdc, szCount);
		m_iCountLabelWidth = sizeText.cx + m_iSideMargin;

		// Button, margin, label, margin, count label
		const auto cx = m_iButtonHeight + m_iSideMargin + iLabelWidth + m_iSideMargin + m_iCountLabelWidth;

		return cx;
	}

	int CountedTextPane::GetFixedHeight()
	{
		auto iHeight = 0;
		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		// Our expand/collapse button
		iHeight += m_iButtonHeight;
		// Control label will be next to this

		if (!m_bCollapsed)
		{
			// Small gap before the edit box
			iHeight += m_iSmallHeightMargin;
		}

		iHeight += m_iSmallHeightMargin; // Bottom margin

		return iHeight;
	}

	int CountedTextPane::GetLines()
	{
		if (m_bCollapsed)
		{
			return 0;
		}

		return LINES_MULTILINEEDIT;
	}

	void CountedTextPane::DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height)
	{
		const auto iVariableHeight = height - GetFixedHeight();
		if (0 != m_paneID)
		{
			y += m_iSmallHeightMargin;
			height -= m_iSmallHeightMargin;
		}

		ViewPane::DeferWindowPos(hWinPosInfo, x, y, width, height);

		if (!m_bCollapsed)
		{
			EC_B_S(m_Count.ShowWindow(SW_SHOW));
			EC_B_S(m_EditBox.ShowWindow(SW_SHOW));

			EC_B_S(::DeferWindowPos(
				hWinPosInfo,
				m_Count.GetSafeHwnd(),
				nullptr,
				x + width - m_iCountLabelWidth,
				y,
				m_iCountLabelWidth,
				m_iLabelHeight,
				SWP_NOZORDER));

			y += m_iLabelHeight + m_iSmallHeightMargin;

			EC_B_S(::DeferWindowPos(
				hWinPosInfo, m_EditBox.GetSafeHwnd(), nullptr, x, y, width, iVariableHeight, SWP_NOZORDER));
		}
		else
		{
			EC_B_S(m_Count.ShowWindow(SW_HIDE));
			EC_B_S(m_EditBox.ShowWindow(SW_HIDE));
		}
	}

	void CountedTextPane::SetCount(size_t iCount) { m_iCount = iCount; }
} // namespace viewpane