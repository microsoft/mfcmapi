#include "stdafx.h"
#include "DropDownPane.h"
#include "String.h"
#include "InterpretProp2.h"
#include <UIFunctions.h>

static wstring CLASS = L"DropDownPane";

DropDownPane* DropDownPane::Create(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, bool bReadOnly)
{
	auto pane = new DropDownPane();
	if (pane)
	{
		pane->SetLabel(uidLabel, bReadOnly);
		pane->Setup(ulDropList, lpuidDropList, nullptr, false);
	}

	return pane;
}

DropDownPane* DropDownPane::CreateArray(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) LPNAME_ARRAY_ENTRY lpnaeDropList, bool bReadOnly)
{
	auto pane = new DropDownPane();
	if (pane)
	{
		pane->Setup(ulDropList, nullptr, lpnaeDropList, false);
		pane->SetLabel(uidLabel, bReadOnly);
	}

	return pane;
}

DropDownPane* DropDownPane::CreateGuid(UINT uidLabel, bool bReadOnly)
{
	auto pane = new DropDownPane();
	if (pane)
	{
		pane->Setup(0, nullptr, nullptr, true);
		pane->SetLabel(uidLabel, bReadOnly);
	}

	return pane;
}

void DropDownPane::Setup(ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, _In_opt_count_(ulDropList) LPNAME_ARRAY_ENTRY lpnaeDropList, bool bGUID)
{
	m_ulDropList = ulDropList;
	m_lpuidDropList = lpuidDropList;
	m_iDropSelection = CB_ERR;
	m_iDropSelectionValue = 0;
	m_lpnaeDropList = lpnaeDropList;
	m_bGUID = bGUID;
}

bool DropDownPane::IsType(__ViewTypes vType)
{
	return CTRL_DROPDOWNPANE == vType;
}

ULONG DropDownPane::GetFlags()
{
	ULONG ulFlags = vpNone;
	if (m_bReadOnly) ulFlags |= vpReadonly;
	return ulFlags;
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
	cxDropDown += ::GetSystemMetrics(SM_CXVSCROLL) + 2 * GetSystemMetrics(SM_CXFIXEDFRAME);

	return max(ViewPane::GetMinWidth(hdc), cxDropDown);
}

int DropDownPane::GetFixedHeight()
{
	auto iHeight = 0;

	if (0 != m_iControl) iHeight += m_iSmallHeightMargin; // Top margin

	if (m_bUseLabelControl)
	{
		iHeight += m_iLabelHeight;
	}

	iHeight += m_iEditHeight; // Height of the dropdown

	// No bottom margin on the DropDown as it's usually tied to the following control
	return iHeight;
}

int DropDownPane::GetLines()
{
	return 0;
}

void DropDownPane::SetWindowPos(int x, int y, int width, int /*height*/)
{
	auto hRes = S_OK;
	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		// height -= m_iSmallHeightMargin;
	}

	if (m_bUseLabelControl)
	{
		EC_B(m_Label.SetWindowPos(
			nullptr,
			x,
			y,
			width,
			m_iLabelHeight,
			SWP_NOZORDER));
		y += m_iLabelHeight;
		// height -= m_iLabelHeight;
	}

	EC_B(m_DropDown.SetWindowPos(NULL, x, y, width, m_iEditHeight, SWP_NOZORDER));
}

void DropDownPane::CreateControl(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	auto hRes = S_OK;

	ViewPane::Initialize(iControl, pParent, hdc);

	auto ulDrops = 1 + (m_ulDropList ? min(m_ulDropList, 4) : 4);
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
		| dwDropStyle,
		CRect(0, 0, 0, dropHeight),
		pParent,
		m_nID));
}

void DropDownPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	CreateControl(iControl, pParent, hdc);

	if (m_lpuidDropList)
	{
		for (ULONG iDropNum = 0; iDropNum < m_ulDropList; iDropNum++)
		{
			auto szDropString = loadstring(m_lpuidDropList[iDropNum]);
			InsertDropString(iDropNum, szDropString, m_lpuidDropList[iDropNum]);
		}
	}
	else if (m_lpnaeDropList)
	{
		for (ULONG iDropNum = 0; iDropNum < m_ulDropList; iDropNum++)
		{
			auto szDropString = wstring(m_lpnaeDropList[iDropNum].lpszName);
			InsertDropString(iDropNum, szDropString, m_lpnaeDropList[iDropNum].ulValue);
		}
	}

	// If this is a GUID list, load up our list of guids
	if (m_bGUID)
	{
		for (ULONG iDropNum = 0; iDropNum < ulPropGuidArray; iDropNum++)
		{
			InsertDropString(iDropNum, GUIDToStringAndName(PropGuidArray[iDropNum].lpGuid), iDropNum);
		}
	}

	m_DropDown.SetCurSel(static_cast<int>(m_iDropSelectionValue));

	m_bInitialized = true;
}

void DropDownPane::InsertDropString(int iRow, _In_ wstring szText, ULONG ulValue)
{
	m_DropDown.InsertString(iRow, wstringTotstring(szText).c_str());
	m_DropDown.SetItemData(iRow, ulValue);
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
	::GetWindowTextW(m_DropDown.m_hWnd, buffer, len);
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
_Check_return_ bool DropDownPane::GetSelectedGUID(bool bByteSwapped, _In_ LPGUID lpSelectedGUID) const
{
	if (!lpSelectedGUID) return NULL;

	auto iCurSel = GetDropDownSelection();
	if (iCurSel != CB_ERR)
	{
		memcpy(lpSelectedGUID, PropGuidArray[iCurSel].lpGuid, sizeof(GUID));
		return true;
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
		memcpy(lpSelectedGUID, lpGUID, sizeof(GUID));
		delete[] lpGUID;
		return true;
	}

	return false;
}

void DropDownPane::SetDropDownSelection(_In_ wstring szText)
{
	auto hRes = S_OK;
	auto text = wstringTotstring(szText).c_str();
	auto iSelect = m_DropDown.SelectString(0, text);

	// if we can't select, try pushing the text in there
	// not all dropdowns will support this!
	if (CB_ERR == iSelect)
	{
		EC_B(::SendMessage(
			m_DropDown.m_hWnd,
			WM_SETTEXT,
			NULL,
			reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(text))));
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