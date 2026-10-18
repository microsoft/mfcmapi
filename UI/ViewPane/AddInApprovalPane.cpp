#include <StdAfx.h>
#include <UI/ViewPane/AddInApprovalPane.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>

namespace viewpane
{
	std::shared_ptr<AddInApprovalPane> AddInApprovalPane::Create(const int paneID)
	{
		auto pane = std::make_shared<AddInApprovalPane>();
		if (pane)
		{
			pane->m_paneID = paneID;
			pane->SetReadOnly(false);
		}

		return pane;
	}

	void AddInApprovalPane::Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);

		const auto buttonText = strings::loadstring(IDS_CLEARADDINAPPROVALS);
		EC_B_S(m_Button.Create(
			strings::wstringTotstring(buttonText).c_str(),
			WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
			CRect(0, 0, 0, 0),
			pParent,
			m_nID));
		m_iButtonWidth = ui::GetTextExtentPoint32(hdc, buttonText).cx + 2 * m_iMargin;
		m_bInitialized = true;
	}

	HDWP AddInApprovalPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		hWinPosInfo = ViewPane::DeferWindowPos(hWinPosInfo, x, y, width, height);
		return ui::DeferWindowPos(
			hWinPosInfo,
			m_Button.GetSafeHwnd(),
			x,
			y + GetHeaderHeight(),
			m_iButtonWidth,
			m_iButtonHeight,
			L"AddInApprovalPane::DeferWindowPos");
	}

	int AddInApprovalPane::GetMinWidth() { return max(ViewPane::GetMinWidth(), m_iButtonWidth); }

	int AddInApprovalPane::GetFixedHeight() { return GetHeaderHeight() + m_iButtonHeight; }

	ULONG AddInApprovalPane::HandleChange(const UINT nID)
	{
		if (nID != m_nID) return ViewPane::HandleChange(nID);

		auto* parent = m_Button.GetParent()->GetParent();
		if (!addin::ClearAddInApprovalCache())
		{
			const auto message = strings::loadstring(IDS_ADDINAPPROVALRESETFAILED);
			parent->MessageBox(strings::wstringTotstring(message).c_str(), _T("MFCMAPI"), MB_OK | MB_ICONERROR);
			return GetID();
		}

		const auto message = strings::loadstring(IDS_ADDINAPPROVALRESETSUCCEEDED);
		parent->MessageBox(strings::wstringTotstring(message).c_str(), _T("MFCMAPI"), MB_OK | MB_ICONINFORMATION);
		return GetID();
	}

	bool AddInApprovalPane::containsWindow(const HWND hWnd) const noexcept
	{
		return m_Button.GetSafeHwnd() == hWnd || ViewPane::containsWindow(hWnd);
	}

	RECT AddInApprovalPane::GetWindowRect() const noexcept
	{
		auto rcButton = RECT{};
		::GetWindowRect(m_Button.GetSafeHwnd(), &rcButton);
		const auto rcHeader = ViewPane::GetWindowRect();
		auto rcPane = RECT{};
		::UnionRect(&rcPane, &rcButton, &rcHeader);
		return rcPane;
	}
} // namespace viewpane
