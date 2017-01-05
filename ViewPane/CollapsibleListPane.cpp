#include "stdafx.h"
#include "CollapsibleListPane.h"

static wstring CLASS = L"CollapsibleListPane";

CollapsibleListPane* CollapsibleListPane::Create(UINT uidLabel, bool bAllowSort, bool bReadOnly, DoListEditCallback callback)
{
	auto pane = new CollapsibleListPane();
	if (pane)
	{
		pane->Setup(bAllowSort, callback);
		pane->SetLabel(uidLabel, bReadOnly);
		pane->m_bCollapsible = true;
	}

	return pane;
}

int CollapsibleListPane::GetFixedHeight()
{
	auto iHeight = 0;
	if (0 != m_iControl) iHeight += m_iSmallHeightMargin; // Top margin

	// Our expand/collapse button
	iHeight += m_iButtonHeight;
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

void CollapsibleListPane::SetWindowPos(int x, int y, int width, int height)
{
	auto hRes = S_OK;
	auto iVariableHeight = height - GetFixedHeight();
	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		height -= m_iSmallHeightMargin;
	}

	ViewPane::SetWindowPos(x, y, width, height);

	y += m_iLabelHeight + m_iSmallHeightMargin;

	if (!m_bCollapsed)
	{
		EC_B(m_List.ShowWindow(SW_SHOW));
		EC_B(m_List.SetWindowPos(
			NULL,
			x,
			y,
			width,
			iVariableHeight,
			SWP_NOZORDER));
		y += iVariableHeight;

		if (!m_bReadOnly)
		{
			// buttons go below the list:
			y += m_iLargeHeightMargin;

			auto iSlotWidth = m_iButtonWidth + m_iMargin;
			auto iOffset = width + m_iSideMargin + m_iMargin;

			for (auto iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
			{
				EC_B(m_ButtonArray[iButton].ShowWindow(SW_SHOW));
				EC_B(m_ButtonArray[iButton].SetWindowPos(
					nullptr,
					iOffset - iSlotWidth * (NUMLISTBUTTONS - iButton),
					y,
					m_iButtonWidth,
					m_iButtonHeight,
					SWP_NOZORDER));
			}
		}
	}
	else
	{
		EC_B(m_List.ShowWindow(SW_HIDE));
		for (auto iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
		{
			EC_B(m_ButtonArray[iButton].ShowWindow(SW_HIDE));
		}
	}
}