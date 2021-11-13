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
			pane->makeCollapsible();
			pane->m_paneID = paneID;
		}

		return pane;
	}

	void CountedTextPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		TextPane::Initialize(pParent, hdc);
		SetCount(m_iCount);
	}

	int CountedTextPane::GetFixedHeight()
	{
		auto iHeight = 0;
		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		iHeight += GetHeaderHeight();

		if (!collapsed())
		{
			// Small gap before the edit box
			iHeight += m_iSmallHeightMargin;
		}

		iHeight += m_iSmallHeightMargin; // Bottom margin

		return iHeight;
	}

	int CountedTextPane::GetLines() { return collapsed() ? 0 : LINES_MULTILINEEDIT; }

	HDWP CountedTextPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		auto curY = y;
		const auto labelHeight = GetHeaderHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		// Layout our label
		hWinPosInfo = EC_D(HDWP, ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y)));

		if (collapsed())
		{
			WC_B_S(m_EditBox.ShowWindow(SW_HIDE));

			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo,
				m_EditBox.GetSafeHwnd(),
				x,
				curY,
				0,
				0,
				L"CountedTextPane::DeferWindowPos::editbox(collapsed)");
		}
		else
		{
			WC_B_S(m_EditBox.ShowWindow(SW_SHOW));

			curY += labelHeight + m_iSmallHeightMargin;

			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo,
				m_EditBox.GetSafeHwnd(),
				x,
				curY,
				width,
				height - (curY - y) - m_iSmallHeightMargin,
				L"CountedTextPane::DeferWindowPos::editbox");
		}

		return hWinPosInfo;
	}

	void CountedTextPane::SetCount(size_t iCount)
	{
		m_iCount = iCount;
		m_Header.SetRightLabel(strings::format(
			L"%ws: 0x%08X = %u",
			m_szCountLabel.c_str(),
			static_cast<int>(m_iCount),
			static_cast<UINT>(m_iCount))); // STRING_OK
	}
} // namespace viewpane