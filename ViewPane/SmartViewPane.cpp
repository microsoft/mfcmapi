#include "stdafx.h"
#include "SmartViewPane.h"
#include "String.h"
#include "SmartView/SmartView.h"

static wstring CLASS = L"SmartViewPane";

ViewPane* CreateSmartViewPane(UINT uidLabel)
{
	return new SmartViewPane(uidLabel);
}

SmartViewPane::SmartViewPane(UINT uidLabel) :DropDownPane(uidLabel, true, ulSmartViewParserTypeArray, nullptr, SmartViewParserTypeArray, false)
{
	m_lpTextPane = static_cast<TextPane*>(CreateMultiLinePane(NULL, true));
	m_bHasData = false;
	m_bDoDropDown = true;
	m_bReadOnly = true;
}

SmartViewPane::~SmartViewPane()
{
	if (m_lpTextPane) delete m_lpTextPane;
}

bool SmartViewPane::IsType(__ViewTypes vType)
{
	return CTRL_SMARTVIEWPANE == vType || DropDownPane::IsType(vType);
}

ULONG SmartViewPane::GetFlags()
{
	return DropDownPane::GetFlags() | vpCollapsible;
}

int SmartViewPane::GetFixedHeight()
{
	if (!m_bDoDropDown && !m_bHasData) return 0;

	auto iHeight = 0;

	if (0 != m_iControl) iHeight += m_iSmallHeightMargin; // Top margin

	// Our expand/collapse button
	iHeight += m_iButtonHeight;
	// Control label will be next to this

	if (m_bDoDropDown && !m_bCollapsed)
	{
		iHeight += m_iEditHeight; // Height of the dropdown
		iHeight += m_lpTextPane->GetFixedHeight();
	}

	return iHeight;
}

int SmartViewPane::GetLines()
{
	auto iStructType = GetDropDownSelectionValue();
	if (!m_bCollapsed && (m_bHasData || iStructType))
	{
		return m_lpTextPane->GetLines();
	}

	return 0;
}

void SmartViewPane::SetWindowPos(int x, int y, int width, int height)
{
	auto hRes = S_OK;
	if (!m_bDoDropDown && !m_bHasData)
	{
		EC_B(m_CollapseButton.ShowWindow(SW_HIDE));
		EC_B(m_Label.ShowWindow(SW_HIDE));
		EC_B(m_DropDown.ShowWindow(SW_HIDE));
		if (m_lpTextPane) m_lpTextPane->ShowWindow(SW_HIDE);
	}
	else
	{
		EC_B(m_CollapseButton.ShowWindow(SW_SHOW));
		EC_B(m_Label.ShowWindow(SW_SHOW));
	}

	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		height -= m_iSmallHeightMargin;
	}

	ViewPane::SetWindowPos(x, y, width, height);

	y += m_iLabelHeight + m_iSmallHeightMargin;
	height -= m_iButtonHeight + m_iSmallHeightMargin;

	if (m_bCollapsed)
	{
		EC_B(m_DropDown.ShowWindow(SW_HIDE));
		if (m_lpTextPane) m_lpTextPane->ShowWindow(SW_HIDE);
	}
	else
	{
		if (m_bDoDropDown)
		{
			EC_B(m_DropDown.ShowWindow(SW_SHOW));
			EC_B(m_DropDown.SetWindowPos(NULL, x, y, width, m_iEditHeight, SWP_NOZORDER));

			y += m_iEditHeight;
			height -= m_iEditHeight;
		}

		if (m_lpTextPane)
		{
			m_lpTextPane->ShowWindow(SW_SHOW);
			m_lpTextPane->SetWindowPos(x, y, width, height);
		}
	}
}

void SmartViewPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	DoInit(iControl, pParent, hdc);

	if (SmartViewParserTypeArray)
	{
		for (ULONG iDropNum = 0; iDropNum < ulSmartViewParserTypeArray; iDropNum++)
		{
			auto szDropString = wstring(SmartViewParserTypeArray[iDropNum].lpszName);
			InsertDropString(iDropNum, szDropString, SmartViewParserTypeArray[iDropNum].ulValue);
		}
	}

	m_DropDown.SetCurSel(static_cast<int>(m_iDropSelectionValue));

	// Passing a control # of 1 gives us a built in margin
	m_lpTextPane->Initialize(1, pParent, hdc);

	m_bInitialized = true;
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
	m_lpTextPane->SetMargins(iMargin, iSideMargin, iLabelHeight, iSmallHeightMargin, iLargeHeightMargin, iButtonHeight, iEditHeight);
	ViewPane::SetMargins(iMargin, iSideMargin, iLabelHeight, iSmallHeightMargin, iLargeHeightMargin, iButtonHeight, iEditHeight);
}

void SmartViewPane::SetStringW(wstring szMsg)
{
	if (!szMsg.empty())
	{
		m_bHasData = true;
	}
	else
	{
		m_bHasData = false;
	}

	m_lpTextPane->SetStringW(szMsg);
}

void SmartViewPane::DisableDropDown()
{
	m_bDoDropDown = false;
}

void SmartViewPane::SetParser(__ParsingTypeEnum iParser)
{
	if (SmartViewParserTypeArray)
	{
		for (ULONG iDropNum = 0; iDropNum < ulSmartViewParserTypeArray; iDropNum++)
		{
			if (iParser == static_cast<__ParsingTypeEnum>(SmartViewParserTypeArray[iDropNum].ulValue))
			{
				SetSelection(iDropNum);
				break;
			}
		}
	}
}

void SmartViewPane::Parse(SBinary myBin)
{
	auto iStructType = static_cast<__ParsingTypeEnum>(GetDropDownSelectionValue());
	auto szSmartView = InterpretBinaryAsString(myBin, iStructType, nullptr);

	m_bHasData = !szSmartView.empty();
	SetStringW(szSmartView);
}