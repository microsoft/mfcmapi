#include "stdafx.h"
#include "..\stdafx.h"
#include "ViewPane.h"
#include "..\UIFunctions.h"
#include "..\String.h"

ViewPane::ViewPane(UINT uidLabel, bool bReadOnly)
{
	m_iControl = -1;
	m_uidLabel = uidLabel;

	m_bCollapsed = false;
	m_bReadOnly = bReadOnly;
	m_bInitialized = false;
	m_iMargin = 0;
	m_iSideMargin = 0;
	m_iLabelHeight = 0;
	m_iSmallHeightMargin = 0;
	m_iLargeHeightMargin = 0;
	m_iButtonHeight = 0;
	m_iEditHeight = 0;
	m_iSmallHeightMargin = 0;
	m_iLabelWidth = 0;

	m_bUseLabelControl = false;
	m_hWndParent = NULL;

	if (m_uidLabel)
	{
		m_bUseLabelControl = true;

		HRESULT hRes = S_OK;
		EC_B(m_szLabel.LoadString(m_uidLabel));
	}
}

ViewPane::~ViewPane()
{
}

bool ViewPane::IsType(__ViewTypes vType)
{
	return CTRL_UNKNOWN == vType;
}

void ViewPane::SetWindowPos(int x, int y, int width, int /*height*/)
{
	HRESULT hRes = S_OK;

	if (vpCollapsible & GetFlags())
	{
		StyleButton(m_CollapseButton.m_hWnd, m_bCollapsed ? bsUpArrow : bsDownArrow);
		m_CollapseButton.SetWindowPos(NULL, x, y, width, m_iLabelHeight, SWP_NOZORDER);
	}

	EC_B(m_Label.SetWindowPos(
		0,
		x + m_iButtonHeight,
		y,
		m_iLabelWidth,
		m_iLabelHeight,
		SWP_NOZORDER));
}

void ViewPane::Initialize(int iControl, _In_ CWnd* pParent, _In_opt_ HDC /*hdc*/)
{
	HRESULT hRes = S_OK;
	m_iControl = iControl;
	if (pParent) m_hWndParent = pParent->m_hWnd;
	UINT iCurIDLabel = IDC_PROP_CONTROL_ID_BASE + 2 * m_iControl;
	m_nID = IDC_PROP_CONTROL_ID_BASE + 2 * m_iControl + 1;

	EC_B(m_Label.Create(
		WS_CHILD
		| WS_CLIPSIBLINGS
		| ES_READONLY
		| WS_VISIBLE,
		CRect(0, 0, 0, 0),
		pParent,
		iCurIDLabel));
	m_Label.SetWindowText(m_szLabel);
	SubclassLabel(m_Label.m_hWnd);

	if (vpCollapsible & GetFlags())
	{
		StyleLabel(m_Label.m_hWnd, lsPaneHeader);

		EC_B(m_CollapseButton.Create(
			NULL,
			WS_TABSTOP
			| WS_CHILD
			| WS_CLIPSIBLINGS
			| WS_VISIBLE,
			CRect(0, 0, 0, 0),
			pParent,
			IDD_COLLAPSE + iControl));
	}
}

void ViewPane::CommitUIValues()
{
}

ULONG ViewPane::GetFlags()
{
	ULONG ulFlags = vpNone;
	if (m_bReadOnly) ulFlags |= vpReadonly;
	return ulFlags;
}

int ViewPane::GetMinWidth(_In_ HDC hdc)
{
	SIZE sizeText = { 0 };
	::GetTextExtentPoint32(hdc, m_szLabel, m_szLabel.GetLength(), &sizeText);
	m_iLabelWidth = sizeText.cx;
	return m_iLabelWidth;
}

ULONG ViewPane::HandleChange(UINT nID)
{
	if ((UINT)(IDD_COLLAPSE + m_iControl) == nID)
	{
		OnToggleCollapse();
		return m_iControl;
	}

	return (ULONG)-1;
}

void ViewPane::OnToggleCollapse()
{
	m_bCollapsed = !m_bCollapsed;

	// Trigger a redraw
	::PostMessage(m_hWndParent, WM_COMMAND, IDD_RECALCLAYOUT, NULL);
}

void ViewPane::SetMargins(
	int iMargin,
	int iSideMargin,
	int iLabelHeight, // Height of the label
	int iSmallHeightMargin,
	int iLargeHeightMargin,
	int iButtonHeight, // Height of buttons below the control
	int iEditHeight) // height of an edit control
{
	m_iMargin = iMargin;
	m_iSideMargin = iSideMargin;
	m_iLabelHeight = iLabelHeight;
	m_iSmallHeightMargin = iSmallHeightMargin;
	m_iLargeHeightMargin = iLargeHeightMargin;
	m_iButtonHeight = iButtonHeight;
	m_iEditHeight = iEditHeight;
}

void ViewPane::SetAddInLabel(_In_z_ LPWSTR szLabel)
{
	m_szLabel = wstringToCString(szLabel);
}

bool ViewPane::MatchID(UINT nID)
{
	return nID == m_nID;
}