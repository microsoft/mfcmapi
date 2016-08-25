#include "stdafx.h"
#include "..\stdafx.h"
#include "CountedTextPane.h"
#include "..\UIFunctions.h"

static wstring CLASS = L"CountedTextPane";

ViewPane* CreateCountedTextPane(UINT uidLabel, bool bReadOnly, UINT uidCountLabel)
{
	CountedTextPane* lpPane = new CountedTextPane(uidLabel, bReadOnly, uidCountLabel);
	return lpPane;
}

CountedTextPane::CountedTextPane(UINT uidLabel, bool bReadOnly, UINT uidCountLabel) :TextPane(uidLabel, bReadOnly, true)
{
	m_iCountLabelWidth = 0;

	m_uidCountLabel = uidCountLabel;
	m_iCount = 0;

	if (m_uidCountLabel)
	{
		m_szCountLabel = loadstring(m_uidCountLabel);
	}
}

bool CountedTextPane::IsType(__ViewTypes vType)
{
	return (CTRL_COUNTEDTEXTPANE == vType) || TextPane::IsType(vType);
}

ULONG CountedTextPane::GetFlags()
{
	return TextPane::GetFlags() | vpCollapsible;
}

int CountedTextPane::GetMinWidth(_In_ HDC hdc)
{
	int cx = 0;
	int iLabelWidth = TextPane::GetMinWidth(hdc);

	wstring szCount = format(L"%ws: 0x%08X = %u", m_szCountLabel.c_str(), (int)m_iCount, (UINT)m_iCount); // STRING_OK
	SetWindowTextW(m_Count.m_hWnd, szCount.c_str());

	SIZE sizeText = { 0 };
	::GetTextExtentPoint32W(hdc, szCount.c_str(), static_cast<int>(szCount.length()), &sizeText);
	m_iCountLabelWidth = sizeText.cx + m_iSideMargin;

	// Button, margin, label, margin, count label
	cx = m_iButtonHeight + m_iSideMargin + iLabelWidth + m_iSideMargin + m_iCountLabelWidth;

	return cx;
}

int CountedTextPane::GetFixedHeight()
{
	int iHeight = 0;
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
	else
	{
		return LINES_MULTILINEEDIT;
	}
}

void CountedTextPane::SetWindowPos(int x, int y, int width, int height)
{
	HRESULT hRes = S_OK;
	int(iVariableHeight) = height - GetFixedHeight();
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
			0,
			x + width - m_iCountLabelWidth,
			y,
			m_iCountLabelWidth,
			m_iLabelHeight,
			SWP_NOZORDER));

		y += m_iLabelHeight + m_iSmallHeightMargin;
		height -= m_iLabelHeight + m_iSmallHeightMargin;

		EC_B(m_EditBox.SetWindowPos(NULL, x, y, width, iVariableHeight, SWP_NOZORDER));
	}
	else
	{
		EC_B(m_Count.ShowWindow(SW_HIDE));
		EC_B(m_EditBox.ShowWindow(SW_HIDE));
	}
	//	height -= m_iSmallHeightMargin; // This is the bottom margin
}

void CountedTextPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	HRESULT hRes = S_OK;

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
	StyleLabel(m_Count.m_hWnd, lsPaneHeader);

	TextPane::Initialize(iControl, pParent, hdc);
}

void CountedTextPane::SetCount(size_t iCount)
{
	m_iCount = iCount;
}