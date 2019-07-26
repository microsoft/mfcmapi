#include <StdAfx.h>
#include <UI/ViewPane/TreePane.h>
#include <utility>
#include <core/utility/output.h>

namespace viewpane
{
	TreePane* TreePane::Create(int paneID, UINT uidLabel, bool bReadOnly)
	{
		auto pane = new (std::nothrow) TreePane();
		if (pane)
		{
			if (uidLabel)
			{
				pane->SetLabel(uidLabel);
				pane->m_bCollapsible = true;
			}

			pane->SetReadOnly(bReadOnly);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	void TreePane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);
		m_Tree.Create(pParent, m_bReadOnly);

		if (InitializeCallback) InitializeCallback(m_Tree);
		m_bInitialized = true;
	}

	int TreePane::GetFixedHeight()
	{
		auto iHeight = 0;
		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		const auto labelHeight = GetLabelHeight();

		if (labelHeight)
		{
			iHeight += labelHeight;
		}

		if (m_bCollapsible && !m_bCollapsed)
		{
			iHeight += m_iSmallHeightMargin;
		}

		iHeight += m_iSmallHeightMargin;

		return iHeight;
	}

	void TreePane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		output::DebugPrint(
			output::DBGDraw, L"TreePane::DeferWindowPos x:%d y:%d width:%d height:%d \n", x, y, width, height);

		auto curY = y;
		const auto labelHeight = GetLabelHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		WC_B_S(m_Tree.ShowWindow(m_bCollapsed ? SW_HIDE : SW_SHOW));
		// Layout our label
		ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y));

		if (labelHeight)
		{
			curY += labelHeight + m_iSmallHeightMargin;
		}

		auto treeHeight = height - (curY - y) - m_iSmallHeightMargin;

		EC_B_S(::DeferWindowPos(hWinPosInfo, m_Tree.GetSafeHwnd(), nullptr, x, curY, width, treeHeight, SWP_NOZORDER));
	}
} // namespace viewpane