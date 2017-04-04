#include "stdafx.h"
#include "DropDownPane.h"
#include "String.h"
#include "InterpretProp2.h"
#include <UIFunctions.h>

static wstring CLASS = L"DropDownPane";

DropDownPane* DropDownPane::Create(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, bool bReadOnly)
{
	auto pane = new DropDownPane();
	if (pane && lpuidDropList)
	{
		for (ULONG iDropNum = 0; iDropNum < ulDropList; iDropNum++)
		{
			pane->InsertDropString(loadstring(lpuidDropList[iDropNum]), lpuidDropList[iDropNum]);
		}
	}

	if (pane)
	{
		pane->SetLabel(uidLabel, bReadOnly);
	}

	return pane;
}

DropDownPane* DropDownPane::CreateGuid(UINT uidLabel, bool bReadOnly)
{
	auto pane = new DropDownPane();
	if (pane)
	{
		for (ULONG iDropNum = 0; iDropNum < PropGuidArray.size(); iDropNum++)
		{
			pane->InsertDropString(GUIDToStringAndName(PropGuidArray[iDropNum].lpGuid), iDropNum);
		}

		pane->SetLabel(uidLabel, bReadOnly);
	}

	return pane;
}

DropDownPane::DropDownPane()
{
	m_iDropSelection = CB_ERR;
	m_iDropSelectionValue = 0;
}

int DropDownPane::GetMinWidth(_In_ HDC hdc)
{
	auto cxDropDown = 0;
	for (auto iDropString = 0; iDropString < m_DropDown.GetCount(); iDropString++)
	{
		auto szDropString = GetLBText(m_DropDown.m_hWnd, iDropString);
		auto sizeDrop = GetTextExtentPoint32(hdc, szDropString);
		cxDropDown = max(cxDropDown, sizeDrop.cx);
	}

	// Add scroll bar and margins for our frame
	cxDropDown += GetSystemMetrics(SM_CXVSCROLL) + 2 * GetSystemMetrics(SM_CXFIXEDFRAME);

	return max(ViewPane::GetMinWidth(hdc), cxDropDown);
}

int DropDownPane::GetFixedHeight()
{
	auto iHeight = 0;

	if (0 != m_iControl) iHeight += m_iSmallHeightMargin; // Top margin

	if (!m_szLabel.empty())
	{
		iHeight += m_iLabelHeight;
	}

	iHeight += m_iEditHeight; // Height of the dropdown

	iHeight += m_iLargeHeightMargin; // Bottom margin

	return iHeight;
}

void DropDownPane::SetWindowPos(int x, int y, int width, int /*height*/)
{
	auto hRes = S_OK;
	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
	}

	if (!m_szLabel.empty())
	{
		EC_B(m_Label.SetWindowPos(
			nullptr,
			x,
			y,
			width,
			m_iLabelHeight,
			SWP_NOZORDER));
		y += m_iLabelHeight;
	}

	// Note - Real height of a combo box is fixed at m_iEditHeight
	// Height we set here influences the amount of dropdown entries we see
	// This will give us something between 4 and 10 entries
	auto ulDrops = static_cast<int>(min(10, 1 + max(m_DropList.size(), 4)));

	EC_B(m_DropDown.SetWindowPos(NULL, x, y, width, m_iEditHeight * ulDrops, SWP_NOZORDER));
}

void DropDownPane::CreateControl(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	auto hRes = S_OK;

	ViewPane::Initialize(iControl, pParent, hdc);

	auto ulDrops = 1 + (m_DropList.size() ? min(m_DropList.size(), 4) : 4);
	auto dropHeight = ulDrops * (pParent ? GetEditHeight(pParent->m_hWnd) : 0x1e);

	// m_bReadOnly means you can't type...
	DWORD dwDropStyle;
	if (m_bReadOnly)
	{
		dwDropStyle = CBS_DROPDOWNLIST; // does not allow typing
	}
	else
	{
		dwDropStyle = CBS_DROPDOWN; // allows typing
	}

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
		| CBS_NOINTEGRALHEIGHT
		| dwDropStyle,
		CRect(0, 0, 0, static_cast<int>(dropHeight)),
		pParent,
		m_nID));

	SendMessage(m_DropDown.m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetSegoeFont()), false);
}

void DropDownPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	CreateControl(iControl, pParent, hdc);

	auto iDropNum = 0;
	for (const auto& drop : m_DropList)
	{
		m_DropDown.InsertString(iDropNum, wstringTotstring(drop.first).c_str());
		m_DropDown.SetItemData(iDropNum, drop.second);
		iDropNum++;
	}

	m_DropDown.SetCurSel(static_cast<int>(m_iDropSelectionValue));

	m_bInitialized = true;
	SetDropDownSelection(m_lpszSelectionString);
}

void DropDownPane::InsertDropString(_In_ const wstring& szText, ULONG ulValue)
{
	m_DropList.push_back({ szText, ulValue });
}

void DropDownPane::CommitUIValues()
{
	m_iDropSelection = GetDropDownSelection();
	m_iDropSelectionValue = GetDropDownSelectionValue();
	m_lpszSelectionString = GetDropStringUseControl();
	m_bInitialized = false; // must be last
}

_Check_return_ wstring DropDownPane::GetDropStringUseControl() const
{
	auto len = m_DropDown.GetWindowTextLength() + 1;
	auto buffer = new WCHAR[len];
	memset(buffer, 0, sizeof(WCHAR)* len);
	GetWindowTextW(m_DropDown.m_hWnd, buffer, len);
	wstring szOut = buffer;
	delete[] buffer;
	return szOut;
}

// This should work whether the editor is active/displayed or not
_Check_return_ int DropDownPane::GetDropDownSelection() const
{
	if (m_bInitialized) return m_DropDown.GetCurSel();

	// In case we're being called after we're done displaying, use the stored value
	return m_iDropSelection;
}

_Check_return_ DWORD_PTR DropDownPane::GetDropDownSelectionValue() const
{
	if (m_bInitialized)
	{
		auto iSel = m_DropDown.GetCurSel();

		if (CB_ERR != iSel)
		{
			return m_DropDown.GetItemData(iSel);
		}
	}
	else
	{
		return m_iDropSelectionValue;
	}

	return 0;
}

_Check_return_ int DropDownPane::GetDropDown() const
{
	return m_iDropSelection;
}

_Check_return_ DWORD_PTR DropDownPane::GetDropDownValue() const
{
	return m_iDropSelectionValue;
}

// This should work whether the editor is active/displayed or not
GUID DropDownPane::GetSelectedGUID(bool bByteSwapped) const
{
	auto iCurSel = GetDropDownSelection();
	if (iCurSel != CB_ERR)
	{
		return *PropGuidArray[iCurSel].lpGuid;
	}

	// no match - need to do a lookup
	wstring szText;
	if (m_bInitialized)
	{
		szText = GetDropStringUseControl();
	}

	if (szText.empty())
	{
		szText = m_lpszSelectionString;
	}

	auto lpGUID = GUIDNameToGUID(szText, bByteSwapped);
	if (lpGUID)
	{
		auto guid = *lpGUID;
		delete[] lpGUID;
		return guid;
	}

	return{ 0 };
}

void DropDownPane::SetDropDownSelection(_In_ const wstring& szText)
{
	m_lpszSelectionString = szText;
	if (!m_bInitialized) return;

	auto hRes = S_OK;
	auto text = wstringTotstring(m_lpszSelectionString);
	auto iSelect = m_DropDown.SelectString(0, text.c_str());

	// if we can't select, try pushing the text in there
	// not all dropdowns will support this!
	if (CB_ERR == iSelect)
	{
		EC_B(::SendMessage(
			m_DropDown.m_hWnd,
			WM_SETTEXT,
			NULL,
			reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(text.c_str()))));
	}
}

void DropDownPane::SetSelection(DWORD_PTR iSelection)
{
	if (!m_bInitialized)
	{
		m_iDropSelectionValue = iSelection;
	}
	else
	{
		m_DropDown.SetCurSel(static_cast<int>(iSelection));
	}
}