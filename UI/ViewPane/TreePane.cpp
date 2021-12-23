#include <StdAfx.h>
#include <UI/ViewPane/TreePane.h>
#include <utility>
#include <core/utility/output.h>
#include <ui/UIFunctions.h>

namespace viewpane
{
	std::shared_ptr<TreePane> TreePane::Create(int paneID, UINT uidLabel, bool bReadOnly)
	{
		auto pane = std::make_shared<TreePane>();
		if (pane)
		{
			if (uidLabel)
			{
				pane->SetLabel(uidLabel);
				pane->makeCollapsible();
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

	// TreePane Layout:
	// Header: GetHeaderHeight
	// Tree:
	//    Tree: variable
	int TreePane::GetFixedHeight() { return GetHeaderHeight(); }

	HDWP TreePane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		output::DebugPrint(
			output::dbgLevel::Draw, L"TreePane::DeferWindowPos x:%d y:%d width:%d height:%d \n", x, y, width, height);

		auto curY = y;

		// Layout our label
		hWinPosInfo = EC_D(HDWP, ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height));
		curY += GetHeaderHeight();

		WC_B_S(m_Tree.ShowWindow(collapsed() ? SW_HIDE : SW_SHOW));
		hWinPosInfo = ui::DeferWindowPos(
			hWinPosInfo, m_Tree.GetSafeHwnd(), x, curY, width, height - (curY - y), L"TreePane::DeferWindowPos::tree");

		return hWinPosInfo;
	}

	bool TreePane::containsWindow(HWND hWnd) const noexcept
	{
		if (m_Tree.GetSafeHwnd() == hWnd) return true;
		return m_Header.containsWindow(hWnd);
	}
} // namespace viewpane