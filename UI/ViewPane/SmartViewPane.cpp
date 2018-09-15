#include <StdAfx.h>
#include <UI/ViewPane/SmartViewPane.h>
#include <Interpret/String.h>
#include <Interpret/SmartView/SmartView.h>

namespace viewpane
{
	static std::wstring CLASS = L"SmartViewPane";

	SmartViewPane* SmartViewPane::Create(UINT uidLabel)
	{
		auto pane = new (std::nothrow) SmartViewPane();
		if (pane)
		{
			pane->SetLabel(uidLabel, true);
			pane->m_bCollapsible = true;
		}

		return pane;
	}

	void SmartViewPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		m_TextPane.SetMultiline();
		m_TextPane.SetLabel(NULL, true);
		m_bReadOnly = true;

		for (const auto& smartViewParserType : SmartViewParserTypeArray)
		{
			InsertDropString(smartViewParserType.lpszName, smartViewParserType.ulValue);
		}

		DropDownPane::Initialize(pParent, hdc);
		// The control id of this text pane doesn't matter, so leave it at 0
		m_TextPane.Initialize(pParent, hdc);

		m_bInitialized = true;
	}

	int SmartViewPane::GetFixedHeight()
	{
		if (!m_bDoDropDown && !m_bHasData) return 0;

		auto iHeight = 0;

		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		// Our expand/collapse button
		iHeight += m_iButtonHeight;
		// Control label will be next to this

		if (m_bDoDropDown && !m_bCollapsed)
		{
			iHeight += m_iEditHeight; // Height of the dropdown
			iHeight += m_TextPane.GetFixedHeight();
		}

		return iHeight;
	}

	int SmartViewPane::GetLines()
	{
		if (!m_bCollapsed && m_bHasData)
		{
			return m_TextPane.GetLines();
		}

		return 0;
	}

	void SmartViewPane::DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height)
	{
		const auto visibility = !m_bDoDropDown && !m_bHasData ? SW_HIDE : SW_SHOW;
		EC_B_S(m_CollapseButton.ShowWindow(visibility));
		EC_B_S(m_Label.ShowWindow(visibility));

		if (0 != m_paneID)
		{
			y += m_iSmallHeightMargin;
			height -= m_iSmallHeightMargin;
		}

		ViewPane::DeferWindowPos(hWinPosInfo, x, y, width, height);

		y += m_iLabelHeight + m_iSmallHeightMargin;
		height -= m_iButtonHeight + m_iSmallHeightMargin;

		if (!m_bCollapsed)
		{
			if (m_bDoDropDown)
			{
				EC_B_S(::DeferWindowPos(
					hWinPosInfo, m_DropDown.GetSafeHwnd(), nullptr, x, y, width, m_iEditHeight * 10, SWP_NOZORDER));

				y += m_iEditHeight;
				height -= m_iEditHeight;
			}

			m_TextPane.DeferWindowPos(hWinPosInfo, x, y, width, height);
		}

		EC_B_S(m_DropDown.ShowWindow(m_bCollapsed ? SW_HIDE : SW_SHOW));
		m_TextPane.ShowWindow(m_bCollapsed || !m_bHasData ? SW_HIDE : SW_SHOW);
	}

	void SmartViewPane::SetMargins(
		int iMargin,
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iSmallHeightMargin,
		int iLargeHeightMargin,
		int iButtonHeight, // Height of buttons below the control
		int iEditHeight) // height of an edit control
	{
		m_TextPane.SetMargins(
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

		m_TextPane.SetStringW(szMsg);
	}

	void SmartViewPane::DisableDropDown() { m_bDoDropDown = false; }

	void SmartViewPane::SetParser(__ParsingTypeEnum iParser)
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

	void SmartViewPane::Parse(SBinary myBin)
	{
		const auto iStructType = static_cast<__ParsingTypeEnum>(GetDropDownSelectionValue());
		auto szSmartView = smartview::InterpretBinaryAsString(myBin, iStructType, nullptr);

		m_bHasData = !szSmartView.empty();
		SetStringW(szSmartView);
	}
} // namespace viewpane