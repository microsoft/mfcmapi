#include <StdAfx.h>
#include <UI/ViewPane/DropDownPane.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <UI/UIFunctions.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>

namespace viewpane
{
	std::shared_ptr<DropDownPane> DropDownPane::Create(
		const int paneID,
		const UINT uidLabel,
		const ULONG ulDropList,
		_In_opt_count_(ulDropList) UINT* lpuidDropList,
		const bool bReadOnly)
	{
		auto pane = std::make_shared<DropDownPane>();
		if (pane)
		{
			if (lpuidDropList)
			{
				for (ULONG iDropNum = 0; iDropNum < ulDropList; iDropNum++)
				{
					pane->InsertDropString(strings::loadstring(lpuidDropList[iDropNum]), lpuidDropList[iDropNum]);
				}
			}

			pane->SetLabel(uidLabel);
			pane->SetReadOnly(bReadOnly);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	std::shared_ptr<DropDownPane> DropDownPane::CreateGuid(const int paneID, const UINT uidLabel, const bool bReadOnly)
	{
		auto pane = std::make_shared<DropDownPane>();
		if (pane)
		{
			for (ULONG iDropNum = 0; iDropNum < PropGuidArray.size(); iDropNum++)
			{
				pane->InsertDropString(guid::GUIDToStringAndName(PropGuidArray[iDropNum].lpGuid), iDropNum);
			}

			pane->SetLabel(uidLabel);
			pane->SetReadOnly(bReadOnly);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	int DropDownPane::GetMinWidth()
	{
		auto cxDropDown = 0;
		const auto hdc = GetDC(m_DropDown.GetSafeHwnd());
		const auto hfontOld = SelectObject(hdc, ui::GetSegoeFont());
		for (auto iDropString = 0; iDropString < m_DropDown.GetCount(); iDropString++)
		{
			const auto szDropString = ui::GetLBText(m_DropDown.m_hWnd, iDropString);
			const auto sizeDrop = ui::GetTextExtentPoint32(hdc, szDropString);
			cxDropDown = max(cxDropDown, sizeDrop.cx);
		}

		static_cast<void>(SelectObject(hdc, hfontOld));
		ReleaseDC(m_DropDown.GetSafeHwnd(), hdc);

		// Add scroll bar and margins for our frame
		cxDropDown += GetSystemMetrics(SM_CXVSCROLL) + 2 * GetSystemMetrics(SM_CXFIXEDFRAME);

		return max(ViewPane::GetMinWidth(), cxDropDown);
	}

	int DropDownPane::GetFixedHeight()
	{
		auto iHeight = 0;

		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		iHeight += GetLabelHeight();

		iHeight += m_iEditHeight; // Height of the dropdown

		iHeight += m_iLargeHeightMargin; // Bottom margin

		return iHeight;
	}

	HDWP DropDownPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int /*height*/)
	{
		auto curY = y;
		const auto labelHeight = GetLabelHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		if (!m_szLabel.empty())
		{
			hWinPosInfo = EC_D(
				HDWP,
				::DeferWindowPos(
					hWinPosInfo, m_Label.GetSafeHwnd(), nullptr, x, curY, width, labelHeight, SWP_NOZORDER));
			curY += labelHeight;
		}

		// Note - Real height of a combo box is fixed at m_iEditHeight
		// Height we set here influences the amount of dropdown entries we see
		// This will give us something between 4 and 10 entries
		const auto ulDrops = static_cast<int>(min(10, 1 + max(m_DropList.size(), 4)));

		hWinPosInfo = EC_D(
			HDWP,
			::DeferWindowPos(
				hWinPosInfo, m_DropDown.GetSafeHwnd(), nullptr, x, curY, width, m_iEditHeight * ulDrops, SWP_NOZORDER));

		return hWinPosInfo;
	}

	void DropDownPane::CreateControl(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);

		const auto ulDrops = 1 + (!m_DropList.empty() ? min(m_DropList.size(), 4) : 4);
		const auto dropHeight = ulDrops * (pParent ? ui::GetEditHeight(pParent->m_hWnd) : 0x1e);

		// m_bReadOnly means you can't type...
		DWORD dwDropStyle = 0;
		if (m_bReadOnly)
		{
			dwDropStyle = CBS_DROPDOWNLIST; // does not allow typing
		}
		else
		{
			dwDropStyle = CBS_DROPDOWN; // allows typing
		}

		EC_B_S(m_DropDown.Create(
			WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_BORDER | WS_VISIBLE | WS_VSCROLL | CBS_OWNERDRAWFIXED |
				CBS_HASSTRINGS | CBS_AUTOHSCROLL | CBS_DISABLENOSCROLL | CBS_NOINTEGRALHEIGHT | dwDropStyle,
			CRect(0, 0, 0, static_cast<int>(dropHeight)),
			pParent,
			m_nID));

		SendMessage(m_DropDown.m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(ui::GetSegoeFont()), false);
	}

	void DropDownPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		CreateControl(pParent, hdc);

		auto iDropNum = 0;
		for (const auto& drop : m_DropList)
		{
			m_DropDown.InsertString(iDropNum, strings::wstringTotstring(drop.first).c_str());
			m_DropDown.SetItemData(iDropNum, drop.second);
			iDropNum++;
		}

		m_DropDown.SetCurSel(static_cast<int>(m_iDropSelectionValue));

		m_bInitialized = true;
		SetDropDownSelection(m_lpszSelectionString);
	}

	void DropDownPane::InsertDropString(_In_ const std::wstring& szText, ULONG ulValue)
	{
		m_DropList.emplace_back(szText, ulValue);
	}

	void DropDownPane::CommitUIValues()
	{
		m_iDropSelection = GetDropDownSelection();
		m_iDropSelectionValue = GetDropDownSelectionValue();
		m_lpszSelectionString = GetDropStringUseControl();
		m_bInitialized = false; // must be last
	}

	_Check_return_ std::wstring DropDownPane::GetDropStringUseControl() const
	{
		const auto len = m_DropDown.GetWindowTextLength() + 1;
		auto out = std::wstring(len, '\0');
		GetWindowTextW(m_DropDown.m_hWnd, const_cast<LPWSTR>(out.c_str()), len);
		return out;
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
			const auto iSel = m_DropDown.GetCurSel();

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

	_Check_return_ int DropDownPane::GetDropDown() const { return m_iDropSelection; }

	_Check_return_ DWORD_PTR DropDownPane::GetDropDownValue() const { return m_iDropSelectionValue; }

	// This should work whether the editor is active/displayed or not
	GUID DropDownPane::GetSelectedGUID(bool bByteSwapped) const
	{
		const auto iCurSel = GetDropDownSelection();
		if (iCurSel != CB_ERR)
		{
			return *PropGuidArray[iCurSel].lpGuid;
		}

		// no match - need to do a lookup
		std::wstring szText;
		if (m_bInitialized)
		{
			szText = GetDropStringUseControl();
		}

		if (szText.empty())
		{
			szText = m_lpszSelectionString;
		}

		return guid::GUIDNameToGUID(szText, bByteSwapped);
	}

	void DropDownPane::SetDropDownSelection(_In_ const std::wstring& szText)
	{
		m_lpszSelectionString = szText;
		if (!m_bInitialized) return;

		auto text = strings::wstringTotstring(m_lpszSelectionString);
		const auto iSelect = m_DropDown.SelectString(0, text.c_str());

		// if we can't select, try pushing the text in there
		// not all dropdowns will support this!
		if (CB_ERR == iSelect)
		{
			EC_B_S(::SendMessage(
				m_DropDown.m_hWnd, WM_SETTEXT, NULL, reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(text.c_str()))));
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
} // namespace viewpane