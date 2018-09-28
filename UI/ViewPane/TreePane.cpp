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
			pane->SetLabel(uidLabel, bReadOnly);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	bool TreePane::IsDirty() { return m_bDirty; }

	void TreePane::Initialize(_In_ CWnd* pParent, _In_ HDC /*hdc*/)
	{
		ViewPane::Initialize(pParent, nullptr);
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

		if (m_bCollapsible)
		{
			// Our expand/collapse button
			iHeight += m_iButtonHeight;
		}
		else if (!m_szLabel.empty())
		{
			iHeight += m_iLabelHeight;
		}

		if (!m_bCollapsed)
		{
			iHeight += m_iSmallHeightMargin;

			if (!m_bReadOnly)
			{
				iHeight += m_iLargeHeightMargin + m_iButtonHeight;
			}
		}

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

	void TreePane::DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height)
	{
		output::DebugPrint(DBGDraw, L"TreePane::DeferWindowPos x:%d y:%d width:%d height:%d \n", x, y, width, height);

		auto iVariableHeight = height - GetFixedHeight();
		if (0 != m_paneID)
		{
			y += m_iSmallHeightMargin;
			height -= m_iSmallHeightMargin;
		}

		EC_B_S(m_Tree.ShowWindow(m_bCollapsed ? SW_HIDE : SW_SHOW));
		ViewPane::DeferWindowPos(hWinPosInfo, x, y, width, height);

		if (m_bCollapsible)
		{
			y += m_iLabelHeight + m_iSmallHeightMargin;
		}
		else
		{
			if (!m_szLabel.empty())
			{
				y += m_iLabelHeight;
				height -= m_iLabelHeight;
			}

			height -= m_iSmallHeightMargin; // This is the bottom margin
		}

		EC_B_S(::DeferWindowPos(
			hWinPosInfo,
			m_Tree.GetSafeHwnd(),
			nullptr,
			x,
			y,
			width,
			m_bCollapsible ? iVariableHeight : height,
			SWP_NOZORDER));
	}

	void TreePane::CommitUIValues() {}
} // namespace viewpane