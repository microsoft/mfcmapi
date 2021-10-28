#include <StdAfx.h>
#include <UI/ViewPane/SmartViewPane.h>
#include <UI/ViewPane/TextPane.h>
#include <core/utility/strings.h>
#include <core/smartview/SmartView.h>
#include <UI/UIFunctions.h>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/addin/mfcmapi.h>
#include <UI/ViewPane/getpane.h>

namespace viewpane
{
	enum __SmartViewPanes
	{
		// Start the SmartViewPane sub panes at 50
		// This means we can only have one SmartViewPane per editor
		SV_TREE = 50,
		SV_TEXT,
		SV_ACTIONBUTTON
	};

	std::shared_ptr<SmartViewPane> SmartViewPane::Create(const int paneID, const UINT uidLabel)
	{
		auto pane = std::make_shared<SmartViewPane>();
		if (pane)
		{
			pane->SetLabel(uidLabel);
			pane->SetReadOnly(true);
			pane->makeCollapsible();
			pane->m_paneID = paneID;
			pane->m_Header.EnableActionButton(SV_ACTIONBUTTON);
		}

		return pane;
	}

	SmartViewPane::SmartViewPane() { m_Splitter = std::make_shared<SplitterPane>(); }

	void SmartViewPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		m_bReadOnly = true;

		for (const auto& smartViewParserType : SmartViewParserTypeArray)
		{
			InsertDropString(smartViewParserType.lpszName, static_cast<ULONG>(smartViewParserType.type));
		}

		DropDownPane::Initialize(pParent, hdc);

		m_TreePane = TreePane::Create(SV_TREE, 0, true);
		m_TreePane->m_Tree.ItemSelectedCallback = [&](auto _1) { return ItemSelected(_1); };
		m_TreePane->m_Tree.OnCustomDrawCallback = [&](auto _1, auto _2, auto _3) { return OnCustomDraw(_1, _2, _3); };

		m_Splitter->SetPaneOne(m_TreePane);
		m_Splitter->SetPaneTwo(TextPane::CreateMultiLinePane(SV_TEXT, 0, true));
		m_Splitter->Initialize(pParent, hdc);

		m_bInitialized = true;
		Parse(m_bins);
	}

	int SmartViewPane::GetFixedHeight()
	{
		if (!m_bDoDropDown && !m_bHasData) return 0;

		auto iHeight = 0;

		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		iHeight += GetHeaderHeight();

		if (m_bDoDropDown && !collapsed())
		{
			iHeight += m_iEditHeight; // Height of the dropdown
			iHeight += m_Splitter->GetFixedHeight();
		}

		return iHeight;
	}

	int SmartViewPane::GetLines()
	{
		if (!collapsed() && m_bHasData)
		{
			return m_Splitter->GetLines();
		}

		return 0;
	}

	ULONG SmartViewPane::HandleChange(const UINT nID)
	{
		if (nID == SV_ACTIONBUTTON)
		{
			HandleAction();
			return nID;
		}

		auto paneID = m_Splitter->HandleChange(nID);
		if (paneID != static_cast<ULONG>(-1)) return paneID;

		return ViewPane::HandleChange(nID);
	}

	void SmartViewPane::HandleAction()
	{
		if (!OnActionButton || m_bins.size() != 1) return;
		const auto bin = SBinary{static_cast<ULONG>(m_bins[0].size()), m_bins[0].data()};
		OnActionButton(bin);
	}

	void SmartViewPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		const auto visibility = !m_bDoDropDown && !m_bHasData ? SW_HIDE : SW_SHOW;
		WC_B_S(m_Header.ShowWindow(visibility));

		auto curY = y;
		const auto labelHeight = GetHeaderHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		// Layout our label
		ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y));

		curY += labelHeight + m_iSmallHeightMargin;

		if (!collapsed())
		{
			if (m_bDoDropDown)
			{
				EC_B_S(::DeferWindowPos(
					hWinPosInfo, m_DropDown.GetSafeHwnd(), nullptr, x, curY, width, m_iEditHeight * 10, SWP_NOZORDER));

				curY += m_iEditHeight;
			}

			m_Splitter->DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y));
		}

		WC_B_S(m_DropDown.ShowWindow(collapsed() ? SW_HIDE : SW_SHOW));
		m_Splitter->ShowWindow(collapsed() || !m_bHasData ? SW_HIDE : SW_SHOW);
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
		m_Splitter->SetMargins(
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

		auto lpPane = std::dynamic_pointer_cast<TextPane>(GetPaneByID(m_Splitter, SV_TEXT));

		if (lpPane)
		{
			lpPane->SetStringW(szMsg);
		}
	}

	void SmartViewPane::SetParser(const parserType parser)
	{
		for (size_t iDropNum = 0; iDropNum < SmartViewParserTypeArray.size(); iDropNum++)
		{
			if (parser == static_cast<parserType>(SmartViewParserTypeArray[iDropNum].type))
			{
				SetSelection(iDropNum);
				break;
			}
		}
	}

	void SmartViewPane::Parse(const std::vector<std::vector<BYTE>>& myBins)
	{
		m_bins = myBins;
		if (!m_bInitialized) return;

		// Clear the visual tree before we recompute treeData so we have nothing pointing to data in treeData
		if (m_TreePane) m_TreePane->m_Tree.Refresh();

		const auto iStructType = static_cast<parserType>(GetDropDownSelectionValue());
		if (OnActionButton && m_bins.size() == 1 && iStructType == parserType::ENTRYID)
		{
			m_Header.SetActionButton(L"Open");
		}
		else
		{
			m_Header.SetActionButton(L"");
		}

		treeData = smartview::block::create(m_bins.size() > 1 ? L"Multivalued Property" : L"");
		auto source = 0;
		for (auto& bin : m_bins)
		{
			auto svp = smartview::InterpretBinary({static_cast<ULONG>(bin.size()), bin.data()}, iStructType, nullptr);
			if (svp)
			{
				svp->setSource(source++);
				if (m_bins.size() == 1)
				{
					treeData = svp;
				}
				else
				{
					treeData->addChild(svp);
				}
			}
		}

		AddChildren(nullptr, treeData);
		SetStringW(treeData->toString());
	}

	void SmartViewPane::AddChildren(HTREEITEM parent, const std::shared_ptr<smartview::block>& data)
	{
		if (!m_TreePane || !data) return;

		auto root = HTREEITEM{};
		// If the node is a header with no text, merge the children up one level
		if (data->getText().empty())
		{
			root = parent;
		}
		else
		{
			// This loans pointers to our blocks to the visual tree without refcounting.
			// Care must be taken to ensure we never release treeData without first clearing the UI.
			root = m_TreePane->m_Tree.AddChildNode(data->getText(), parent, data.get(), nullptr);
		}

		for (const auto& item : data->getChildren())
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
		const auto pane = std::dynamic_pointer_cast<TreePane>(GetPaneByID(m_Splitter, SV_TREE));
		if (!pane) return;

		auto tvi = TVITEM{};
		tvi.mask = TVIF_PARAM;
		tvi.hItem = hItem;
		::SendMessage(pane->m_Tree.GetSafeHwnd(), TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&tvi));
		const auto lpData = reinterpret_cast<smartview::block*>(tvi.lParam);
		if (lpData)
		{
			if (registry::uiDiag)
			{
				SetStringW(lpData->toString());
			}

			if (OnItemSelected)
			{
				OnItemSelected(lpData);
			}
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
			if (!registry::uiDiag) return;
			const auto hItem = reinterpret_cast<HTREEITEM>(lvcd->nmcd.dwItemSpec);
			if (hItem)
			{
				auto tvi = TVITEM{};
				tvi.mask = TVIF_PARAM;
				tvi.hItem = hItem;
				TreeView_GetItem(lvcd->nmcd.hdr.hwndFrom, &tvi);
				const auto lpData = reinterpret_cast<smartview::block*>(tvi.lParam);
				if (lpData && !lpData->isHeader())
				{
					const auto bin = strings::BinToHexString(
						m_bins[lpData->getSource()].data() + lpData->getOffset(), lpData->getSize(), false);

					const auto blockString =
						strings::format(L"(%d, %d) %ws", lpData->getOffset(), lpData->getSize(), bin.c_str());
					const auto size = ui::GetTextExtentPoint32(lvcd->nmcd.hdc, blockString);
					auto rect = RECT{};
					TreeView_GetItemRect(lvcd->nmcd.hdr.hwndFrom, hItem, &rect, 1);
					rect.left = rect.right;
					rect.right += size.cx;
					ui::DrawSegoeTextW(
						lvcd->nmcd.hdc,
						blockString,
						MyGetSysColor(ui::uiColor::Glow),
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