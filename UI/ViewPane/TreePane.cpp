#include <StdAfx.h>
#include <UI/ViewPane/TreePane.h>
#include <UI/UIFunctions.h>
#include <utility>

namespace viewpane
{
	static std::wstring CLASS = L"TreePane";

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

	bool TreePane::IsDirty() { return m_bDirty; }

	void TreePane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);
		m_Tree.Create(pParent, 0);

		// TODO: Don't leave this here
		TVINSERTSTRUCTW tvInsert = {nullptr};
		auto szName = std::wstring(L"TEST");
		tvInsert.hParent = nullptr;
		tvInsert.hInsertAfter = TVI_SORT;
		tvInsert.item.mask = TVIF_CHILDREN | TVIF_TEXT;
		tvInsert.item.cChildren = I_CHILDRENCALLBACK;
		tvInsert.item.pszText = const_cast<LPWSTR>(szName.c_str());
		::SendMessage(m_Tree.GetSafeHwnd(), TVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&tvInsert));

		m_bInitialized = true;
	}

	int TreePane::GetMinWidth(_In_ HDC /*hdc*/)
	{
		return 100;
		//return max(
		//	ViewPane::GetMinWidth(hdc), (int) (NUMLISTBUTTONS * m_iButtonWidth + m_iMargin * (NUMLISTBUTTONS - 1)));
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

	int TreePane::GetLines()
	{
		if (m_bCollapsed)
		{
			return 0;
		}

		return 4;
	}

	ULONG TreePane::HandleChange(UINT nID)
	{
		//switch (nID)
		//{
		//}

		return ViewPane::HandleChange(nID);
	}

	void TreePane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		output::DebugPrint(DBGDraw, L"TreePane::DeferWindowPos x:%d y:%d width:%d height:%d \n", x, y, width, height);

		auto curY = y;
		const auto labelHeight = GetLabelHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		EC_B_S(m_Tree.ShowWindow(m_bCollapsed ? SW_HIDE : SW_SHOW));
		// Layout our label
		ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y));

		if (labelHeight)
		{
			curY += labelHeight + m_iSmallHeightMargin;
		}

		auto treeHeight = height - (curY - y) - m_iSmallHeightMargin;

		EC_B_S(::DeferWindowPos(hWinPosInfo, m_Tree.GetSafeHwnd(), nullptr, x, curY, width, treeHeight, SWP_NOZORDER));
	}

	void TreePane::CommitUIValues() {}
} // namespace viewpane