#include "stdafx.h"
#include "..\stdafx.h"
#include "SmartViewPane.h"
#include "..\String.h"
#include "..\UIFunctions.h"
#include "..\SmartView\SmartView.h"

static TCHAR* CLASS = _T("SmartViewPane");

ViewPane* CreateSmartViewPane(UINT uidLabel)
{
	return new SmartViewPane(uidLabel);
}

SmartViewPane::SmartViewPane(UINT uidLabel) :DropDownPane(uidLabel, true, g_cuidParsingTypes, NULL, g_uidParsingTypes, false)
{
	m_lpTextPane = (TextPane*)CreateMultiLinePane(NULL, NULL, true);
	m_bHasData = false;
	m_bDoDropDown = true;
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

	int iHeight = 0;

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
	DWORD_PTR iStructType = GetDropDownSelectionValue();
	if (!m_bCollapsed && (m_bHasData || iStructType))
	{
		return m_lpTextPane->GetLines();
	}
	else
	{
		return 0;
	}
}

void SmartViewPane::SetWindowPos(int x, int y, int width, int height)
{
	HRESULT hRes = S_OK;
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
			// Note - Real height of a combo box is fixed at m_iEditHeight
			// Height we set here influences the amount of dropdown entries we see
			// Only really matters on Win2k and below.
			ULONG ulDrops = 1 + min(g_cuidParsingTypes, 4);
			EC_B(m_DropDown.SetWindowPos(NULL, x, y, width, m_iEditHeight * ulDrops, SWP_NOZORDER));

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
	HRESULT hRes = S_OK;

	ViewPane::Initialize(iControl, pParent, hdc);

	EC_B(m_DropDown.Create(
		WS_TABSTOP
		| WS_CHILD
		| WS_CLIPSIBLINGS
		| WS_BORDER
		| WS_VISIBLE
		| WS_VSCROLL
		| CBS_OWNERDRAWFIXED
		| CBS_HASSTRINGS
		| CBS_AUTOHSCROLL
		| CBS_DISABLENOSCROLL
		| CBS_DROPDOWNLIST, // does not allow typing
		CRect(0, 0, 0, 0),
		pParent,
		m_nID));

	ULONG iDropNum = 0;

	if (g_uidParsingTypes)
	{
		for (iDropNum = 0; iDropNum < g_cuidParsingTypes; iDropNum++)
		{
#ifdef UNICODE
			m_DropDown.InsertString(
				iDropNum,
				g_uidParsingTypes[iDropNum].lpszName);
#else
			LPSTR szAnsiName = NULL;
			EC_H(UnicodeToAnsi(g_uidParsingTypes[iDropNum].lpszName, &szAnsiName));
			if (SUCCEEDED(hRes))
			{
				m_DropDown.InsertString(
					iDropNum,
					szAnsiName);
			}
			delete[] szAnsiName;
#endif
			m_DropDown.SetItemData(
				iDropNum,
				g_uidParsingTypes[iDropNum].ulValue);
		}
	}
	m_DropDown.SetCurSel((int)m_iDropSelectionValue);

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

void SmartViewPane::SetStringW(_In_opt_z_ LPCWSTR szMsg)
{
	if (szMsg && szMsg[0])
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
	ULONG iDropNum = 0;

	if (g_uidParsingTypes)
	{
		for (iDropNum = 0; iDropNum < g_cuidParsingTypes; iDropNum++)
		{
			if (iParser == (__ParsingTypeEnum)g_uidParsingTypes[iDropNum].ulValue)
			{
				SetSelection(iDropNum);
				break;
			}
		}
	}
}

void SmartViewPane::Parse(SBinary myBin)
{
	DWORD_PTR iStructType = GetDropDownSelectionValue();
	wstring szSmartView = InterpretBinaryAsString(myBin, iStructType, NULL, NULL);

	if (!szSmartView.empty())
	{
		m_bHasData = true;
		SetStringW(szSmartView.c_str());
	}
	else
	{
		m_bHasData = false;
		SetStringW(L"");
	}

}