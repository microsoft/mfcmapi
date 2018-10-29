#include <StdAfx.h>
#include <UI/ViewPane/SmartViewPane.h>
#include <UI/ViewPane/TextPane.h>
#include <Interpret/String.h>
#include <Interpret/SmartView/SmartView.h>
#include <UI/UIFunctions.h>

namespace viewpane
{
	enum __SmartViewFields
	{
		SV_TREE,
		SV_TEXT
	};

	SmartViewPane* SmartViewPane::Create(const int paneID, const UINT uidLabel)
	{
		auto pane = new (std::nothrow) SmartViewPane();
		if (pane)
		{
			pane->SetLabel(uidLabel);
			pane->SetReadOnly(true);
			pane->m_bCollapsible = true;
			pane->m_paneID = paneID;
		}

		return pane;
	}

	void SmartViewPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		m_bReadOnly = true;

		for (const auto& smartViewParserType : SmartViewParserTypeArray)
		{
			InsertDropString(smartViewParserType.lpszName, smartViewParserType.ulValue);
		}

		DropDownPane::Initialize(pParent, hdc);

		m_TreePane = TreePane::Create(SV_TREE, 0, true);
		m_TreePane->m_Tree.ItemSelectedCallback = [&](auto _1) { return ItemSelected(_1); };
		m_TreePane->m_Tree.OnCustomDrawCallback = [&](auto _1, auto _2, auto _3) { return OnCustomDraw(_1, _2, _3); };

		m_Splitter.SetPaneOne(m_TreePane);
		m_Splitter.SetPaneTwo(TextPane::CreateMultiLinePane(SV_TEXT, 0, true));
		m_Splitter.Initialize(pParent, hdc);

		m_bInitialized = true;
	}

	int SmartViewPane::GetFixedHeight()
	{
		if (!m_bDoDropDown && !m_bHasData) return 0;

		auto iHeight = 0;

		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		iHeight += GetLabelHeight();

		if (m_bDoDropDown && !m_bCollapsed)
		{
			iHeight += m_iEditHeight; // Height of the dropdown
			iHeight += m_Splitter.GetFixedHeight();
		}

		return iHeight;
	}

	int SmartViewPane::GetLines()
	{
		if (!m_bCollapsed && m_bHasData)
		{
			return m_Splitter.GetLines();
		}

		return 0;
	}

	void SmartViewPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		const auto visibility = !m_bDoDropDown && !m_bHasData ? SW_HIDE : SW_SHOW;
		EC_B_S(m_CollapseButton.ShowWindow(visibility));
		EC_B_S(m_Label.ShowWindow(visibility));

		auto curY = y;
		const auto labelHeight = GetLabelHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		// Layout our label
		ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y));

		curY += labelHeight + m_iSmallHeightMargin;

		if (!m_bCollapsed)
		{
			if (m_bDoDropDown)
			{
				EC_B_S(::DeferWindowPos(
					hWinPosInfo, m_DropDown.GetSafeHwnd(), nullptr, x, curY, width, m_iEditHeight * 10, SWP_NOZORDER));

				curY += m_iEditHeight;
			}

			m_Splitter.DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y));
		}

		EC_B_S(m_DropDown.ShowWindow(m_bCollapsed ? SW_HIDE : SW_SHOW));
		m_Splitter.ShowWindow(m_bCollapsed || !m_bHasData ? SW_HIDE : SW_SHOW);
	}

	void SmartViewPane::SetMargins(
		const int iMargin,
		const int iSideMargin,
		const int iLabelHeight, // Height of the label
		const int iSmallHeightMargin,
		const int iLargeHeightMargin,
		const int iButtonHeight, // Height of buttons below the control
		const int iEditHeight) // height of an edit control
	{
		m_Splitter.SetMargins(
			iMargin, iSideMargin, iLabelHeight, iSmallHeightMargin, iLargeHeightMargin, iButtonHeight, iEditHeight);
		ViewPane::SetMargins(
			iMargin, iSideMargin, iLabelHeight, iSmallHeightMargin, iLargeHeightMargin, iButtonHeight, iEditHeight);
	}

	void SmartViewPane::SetStringW(const std::wstring& szMsg)
	{
		if (!szMsg.empty())
		{
			m_bHasData = true;
		}
		else
		{
			m_bHasData = false;
		}

		auto lpPane = dynamic_cast<TextPane*>(m_Splitter.GetPaneByID(SV_TEXT));

		if (lpPane)
		{
			lpPane->SetStringW(szMsg);
		}
	}

	void SmartViewPane::DisableDropDown() { m_bDoDropDown = false; }

	void SmartViewPane::SetParser(const __ParsingTypeEnum iParser)
	{
		for (size_t iDropNum = 0; iDropNum < SmartViewParserTypeArray.size(); iDropNum++)
		{
			if (iParser == static_cast<__ParsingTypeEnum>(SmartViewParserTypeArray[iDropNum].ulValue))
			{
				SetSelection(iDropNum);
				break;
			}
		}
	}

	void SmartViewPane::Parse(const std::vector<BYTE>& myBin)
	{
		m_bin = myBin;
		const auto iStructType = static_cast<__ParsingTypeEnum>(GetDropDownSelectionValue());
		auto szSmartView = std::wstring{};
		auto svp = smartview::GetSmartViewParser(iStructType, nullptr);
		if (svp)
		{
			svp->init(m_bin.size(), m_bin.data());
			szSmartView = svp->ToString();
			treeData = svp->getBlock();
			RefreshTree();
			delete svp;
		}

		if (szSmartView.empty())
		{
			szSmartView =
				smartview::InterpretBinaryAsString(SBinary{ULONG(m_bin.size()), m_bin.data()}, iStructType, nullptr);
		}

		m_bHasData = !szSmartView.empty();
		SetStringW(szSmartView);
	}

	void SmartViewPane::RefreshTree()
	{
		if (!m_TreePane) return;
		m_TreePane->m_Tree.Refresh();

		AddChildren(nullptr, treeData);
	}

	void SmartViewPane::AddChildren(HTREEITEM parent, const smartview::block& data)
	{
		if (!m_TreePane) return;

		const auto root =
			m_TreePane->m_Tree.AddChildNode(data.getText(), parent, reinterpret_cast<LPARAM>(&data), nullptr);
		for (const auto& item : data.getChildren())
		{
			AddChildren(root, item);
		}

		WC_B_S(::SendMessage(
			m_TreePane->m_Tree.GetSafeHwnd(),
			TVM_EXPAND,
			static_cast<WPARAM>(TVE_EXPAND),
			reinterpret_cast<LPARAM>(root)));
	}

	void SmartViewPane::ItemSelected(HTREEITEM hItem)
	{
		const auto pane = dynamic_cast<TreePane*>(m_Splitter.GetPaneByID(SV_TREE));
		if (!pane) return;

		auto tvi = TVITEM{};
		tvi.mask = TVIF_PARAM;
		tvi.hItem = hItem;
		::SendMessage(pane->m_Tree.GetSafeHwnd(), TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&tvi));
		const auto lpData = reinterpret_cast<smartview::block*>(tvi.lParam);
		if (lpData)
		{
			SetStringW(lpData->ToString());
		}
	}

	void
	SmartViewPane::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* /*pResult*/, _In_ HTREEITEM /*hItemCurHover*/) const
	{
		const auto lvcd = reinterpret_cast<LPNMTVCUSTOMDRAW>(pNMHDR);
		if (!lvcd) return;

		switch (lvcd->nmcd.dwDrawStage)
		{
		case CDDS_ITEMPOSTPAINT:
		{
			const auto hItem = reinterpret_cast<HTREEITEM>(lvcd->nmcd.dwItemSpec);
			if (hItem)
			{
				auto tvi = TVITEM{};
				tvi.mask = TVIF_PARAM;
				tvi.hItem = hItem;
				TreeView_GetItem(lvcd->nmcd.hdr.hwndFrom, &tvi);
				const auto lpData = reinterpret_cast<smartview::block*>(tvi.lParam);
				if (lpData)
				{
					const auto blockString = strings::format(L"(%d, %d)", lpData->getOffset(), lpData->getSize());
					const auto size = ui::GetTextExtentPoint32(lvcd->nmcd.hdc, blockString);
					auto rect = RECT{};
					TreeView_GetItemRect(lvcd->nmcd.hdr.hwndFrom, hItem, &rect, 1);
					rect.left = rect.right;
					rect.right += size.cx;
					ui::DrawSegoeTextW(
						lvcd->nmcd.hdc,
						blockString,
						MyGetSysColor(ui::cGlow),
						rect,
						false,
						DT_SINGLELINE | DT_VCENTER | DT_CENTER);
				}
			}
			break;
		}
		}
	}
} // namespace viewpane