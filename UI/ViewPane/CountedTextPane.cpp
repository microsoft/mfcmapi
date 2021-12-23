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

	int CountedTextPane::GetLines() { return collapsed() ? 0 : LINES_MULTILINEEDIT; }

	// CountedTextPane Layout:
	// Header: GetHeaderHeight
	// Collapsible:
	//    Edit box: variable
	int CountedTextPane::GetFixedHeight() { return GetHeaderHeight(); }

	HDWP CountedTextPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		auto curY = y;

		// Layout our label
		hWinPosInfo = EC_D(HDWP, ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height));
		curY += GetHeaderHeight();

		if (collapsed())
		{
			WC_B_S(m_EditBox.ShowWindow(SW_HIDE));
		}
		else
		{
			WC_B_S(m_EditBox.ShowWindow(SW_SHOW));

			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo,
				m_EditBox.GetSafeHwnd(),
				x,
				curY,
				width,
				height - (curY - y),
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

	bool CountedTextPane::containsWindow(HWND hWnd) const noexcept { return TextPane::containsWindow(hWnd); }
} // namespace viewpane