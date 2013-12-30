#include "stdafx.h"
#include "..\stdafx.h"
#include "DropDownPane.h"
#include "..\MAPIFunctions.h"
#include "..\InterpretProp.h"
#include "..\InterpretProp2.h"

static TCHAR* CLASS = _T("DropDownPane");

ViewPane* CreateDropDownPane(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, bool bReadOnly)
{
	return new DropDownPane(uidLabel, bReadOnly, ulDropList, lpuidDropList, NULL, false);
}

ViewPane* CreateDropDownArrayPane(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) LPNAME_ARRAY_ENTRY lpnaeDropList, bool bReadOnly)
{
	return new DropDownPane(uidLabel, bReadOnly, ulDropList, NULL, lpnaeDropList, false);
}

ViewPane* CreateDropDownGuidPane(UINT uidLabel, bool bReadOnly)
{
	return new DropDownPane(uidLabel, bReadOnly, 0, NULL, NULL, true);
}

DropDownPane::DropDownPane(UINT uidLabel, bool bReadOnly, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, _In_opt_count_(ulDropList) LPNAME_ARRAY_ENTRY lpnaeDropList, bool bGUID):ViewPane(uidLabel, bReadOnly)
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
	int cxDropDown = 0;
	ULONG iDropString = 0;
	for (iDropString = 0; iDropString < m_ulDropList; iDropString++)
	{
		SIZE sizeDrop = {0};
		CString szDropString;
		m_DropDown.GetLBText(iDropString,szDropString);
		::GetTextExtentPoint32(hdc, szDropString, szDropString.GetLength(), &sizeDrop);
		cxDropDown = max(cxDropDown, sizeDrop.cx);
	}

	// Add scroll bar and margins for our frame
	cxDropDown += ::GetSystemMetrics(SM_CXVSCROLL) + 2 * GetSystemMetrics(SM_CXFIXEDFRAME);

	return max(ViewPane::GetMinWidth(hdc), cxDropDown);
}

int DropDownPane::GetFixedHeight()
{
	int iHeight = 0;

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

void DropDownPane::SetWindowPos(int x, int y, int width, int height)
{
	HRESULT hRes = S_OK;
	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		height -= m_iSmallHeightMargin;
	}

	if (m_bUseLabelControl)
	{
		EC_B(m_Label.SetWindowPos(
			0,
			x,
			y,
			width,
			m_iLabelHeight,
			SWP_NOZORDER));
		y += m_iLabelHeight;
		height -= m_iLabelHeight;
	}

	// Note - Real height of a combo box is fixed at m_iEditHeight
	// Height we set here influences the amount of dropdown entries we see
	// Only really matters on Win2k and below.
	ULONG ulDrops = 1 + min(m_ulDropList,4);

	EC_B(m_DropDown.SetWindowPos(NULL, x, y, width, m_iEditHeight * ulDrops, SWP_NOZORDER));
}

void DropDownPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC /*hdc*/)
{
	ViewPane::Initialize(iControl, pParent, NULL);

	HRESULT hRes = S_OK;
	// bReadOnly means you can't type...
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
		CRect(0,0,0,0),
		pParent,
		m_nID));

	ULONG iDropNum = 0;
	if (m_lpuidDropList)
	{
		for (iDropNum=0 ; iDropNum < m_ulDropList ; iDropNum++)
		{
			CString szDropString;
			EC_B(szDropString.LoadString(m_lpuidDropList[iDropNum]));
			m_DropDown.InsertString(
				iDropNum,
				szDropString);
			m_DropDown.SetItemData(
				iDropNum,
				m_lpuidDropList[iDropNum]);
		}
	}
	else if (m_lpnaeDropList)
	{
		for (iDropNum=0 ; iDropNum < m_ulDropList ; iDropNum++)
		{
#ifdef UNICODE
			m_DropDown.InsertString(
				iDropNum,
				m_lpnaeDropList[iDropNum].lpszName);
#else
			LPSTR szAnsiName = NULL;
			EC_H(UnicodeToAnsi(m_lpnaeDropList[iDropNum].lpszName,&szAnsiName));
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
				m_lpnaeDropList[iDropNum].ulValue);
		}
	}

	// If this is a GUID list, load up our list of guids
	if (m_bGUID)
	{
		for (iDropNum=0 ; iDropNum < ulPropGuidArray ; iDropNum++)
		{
			LPTSTR szGUID = GUIDToStringAndName(PropGuidArray[iDropNum].lpGuid);
			InsertDropString(iDropNum, szGUID);
			delete[] szGUID;
		}
	}

	m_DropDown.SetCurSel((int) m_iDropSelectionValue);

	m_bInitialized = true;
}

void DropDownPane::InsertDropString(int iRow, _In_z_ LPCTSTR szText)
{
	m_DropDown.InsertString(iRow,szText);
	m_ulDropList++;
}

void DropDownPane::CommitUIValues()
{
	m_iDropSelection = GetDropDownSelection();
	m_iDropSelectionValue = GetDropDownSelectionValue();
	m_lpszSelectionString = GetDropStringUseControl();
	m_bInitialized = false; // must be last
}

_Check_return_ CString DropDownPane::GetDropStringUseControl()
{
	CString szText;
	m_DropDown.GetWindowText(szText);

	return szText;
}

// This should work whether the editor is active/displayed or not
_Check_return_ int DropDownPane::GetDropDownSelection()
{
	if (m_bInitialized) return m_DropDown.GetCurSel();

	// In case we're being called after we're done displaying, use the stored value
	return m_iDropSelection;
}

_Check_return_ DWORD_PTR DropDownPane::GetDropDownSelectionValue()
{
	if (m_bInitialized)
	{
		int iSel = m_DropDown.GetCurSel();

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

_Check_return_ int DropDownPane::GetDropDown()
{
	return m_iDropSelection;
}

_Check_return_ DWORD_PTR DropDownPane::GetDropDownValue()
{
	return m_iDropSelectionValue;
}

// This should work whether the editor is active/displayed or not
_Check_return_ bool DropDownPane::GetSelectedGUID(bool bByteSwapped, _In_ LPGUID lpSelectedGUID)
{
	if (!lpSelectedGUID) return NULL;

	LPCGUID lpGUID = NULL;
	int iCurSel = GetDropDownSelection();
	if (iCurSel != CB_ERR)
	{
		lpGUID = PropGuidArray[iCurSel].lpGuid;
	}
	else
	{
		// no match - need to do a lookup
		CString szText;
		GUID guid = {0};
		szText = GetDropStringUseControl();
		if (szText.IsEmpty()) szText = m_lpszSelectionString;

		// try the GUID like PS_* first
		GUIDNameToGUID((LPCTSTR) szText, &lpGUID);
		if (!lpGUID) // no match - try it like a guid {}
		{
			HRESULT hRes = S_OK;
			WC_H(StringToGUID((LPCTSTR) szText, bByteSwapped, &guid));

			if (SUCCEEDED(hRes))
			{
				lpGUID = &guid;
			}
		}
	}

	if (lpGUID)
	{
		memcpy(lpSelectedGUID,lpGUID,sizeof(GUID));
		return true;
	}
	return false;
}

void DropDownPane::SetDropDownSelection(_In_opt_z_ LPCTSTR szText)
{
	HRESULT hRes = S_OK;

	int iSelect = m_DropDown.SelectString(0,szText);

	// if we can't select, try pushing the text in there
	// not all dropdowns will support this!
	if (CB_ERR == iSelect)
	{
		EC_B(::SendMessage(
			m_DropDown.m_hWnd,
			WM_SETTEXT,
			NULL,
			(LPARAM) szText));
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
		m_DropDown.SetCurSel((int) iSelection);
	}
}