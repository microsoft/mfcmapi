#include "stdafx.h"
#include "CountedTextPane.h"
#include "UIFunctions.h"

static wstring CLASS = L"CountedTextPane";

CountedTextPane* CountedTextPane::Create(UINT uidLabel, bool bReadOnly, UINT uidCountLabel)
{
	auto lpPane = new CountedTextPane();
	if (lpPane)
	{
		lpPane->m_szCountLabel = loadstring(uidCountLabel);
		lpPane->SetMultiline();
		lpPane->SetLabel(uidLabel, bReadOnly);
		lpPane->m_bCollapsible = true;
	}

	return lpPane;
}

CountedTextPane::CountedTextPane()
{
	m_iCountLabelWidth = 0;
	m_iCount = 0;
}

void CountedTextPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	auto hRes = S_OK;

	EC_B(m_Count.Create(
		WS_CHILD
		| WS_CLIPSIBLINGS
		| ES_READONLY
		| WS_VISIBLE,
		CRect(0, 0, 0, 0),
		pParent,
		IDD_COUNTLABEL));
	SetWindowTextW(m_Count.m_hWnd, m_szCountLabel.c_str());
	SubclassLabel(m_Count.m_hWnd);
	StyleLabel(m_Count.m_hWnd, lsPaneHeaderText);

	TextPane::Initialize(iControl, pParent, hdc);
}


int CountedTextPane::GetMinWidth(_In_ HDC hdc)
{
	auto iLabelWidth = TextPane::GetMinWidth(hdc);

	auto szCount = format(L"%ws: 0x%08X = %u", m_szCountLabel.c_str(), static_cast<int>(m_iCount), static_cast<UINT>(m_iCount)); // STRING_OK
	SetWindowTextW(m_Count.m_hWnd, szCount.c_str());

	auto sizeText = GetTextExtentPoint32(hdc, szCount);
	m_iCountLabelWidth = sizeText.cx + m_iSideMargin;

	// Button, margin, label, margin, count label
	auto cx = m_iButtonHeight + m_iSideMargin + iLabelWidth + m_iSideMargin + m_iCountLabelWidth;

	return cx;
}

int CountedTextPane::GetFixedHeight()
{
	auto iHeight = 0;
	if (0 != m_iControl) iHeight += m_iSmallHeightMargin; // Top margin

	// Our expand/collapse button
	iHeight += m_iButtonHeight;
	// Control label will be next to this

	if (!m_bCollapsed)
	{
		// Small gap before the edit box
		iHeight += m_iSmallHeightMargin;
	}

	iHeight += m_iSmallHeightMargin; // Bottom margin

	return iHeight;
}

int CountedTextPane::GetLines()
{
	if (m_bCollapsed)
	{
		return 0;
	}

	return LINES_MULTILINEEDIT;
}

void CountedTextPane::SetWindowPos(int x, int y, int width, int height)
{
	auto hRes = S_OK;
	auto iVariableHeight = height - GetFixedHeight();
	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		height -= m_iSmallHeightMargin;
	}

	ViewPane::SetWindowPos(x, y, width, height);

	if (!m_bCollapsed)
	{
		EC_B(m_Count.ShowWindow(SW_SHOW));
		EC_B(m_EditBox.ShowWindow(SW_SHOW));

		EC_B(m_Count.SetWindowPos(
			nullptr,
			x + width - m_iCountLabelWidth,
			y,
			m_iCountLabelWidth,
			m_iLabelHeight,
			SWP_NOZORDER));

		y += m_iLabelHeight + m_iSmallHeightMargin;

		EC_B(m_EditBox.SetWindowPos(NULL, x, y, width, iVariableHeight, SWP_NOZORDER));
	}
	else
	{
		EC_B(m_Count.ShowWindow(SW_HIDE));
		EC_B(m_EditBox.ShowWindow(SW_HIDE));
	}
}

void CountedTextPane::SetCount(size_t iCount)
{
	m_iCount = iCount;
}