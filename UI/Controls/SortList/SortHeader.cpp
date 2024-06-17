#include <StdAfx.h>
#include <UI/Controls/SortList/SortHeader.h>
#include <UI/UIFunctions.h>
#include <core/utility/output.h>

namespace controls::sortlistctrl
{
	CSortHeader::CSortHeader() : m_bTooltipDisplayed(false)
	{
		m_hwndTip = nullptr;
		m_hwndParent = nullptr;
		ZeroMemory(&m_ti, sizeof(TOOLINFO));
	}

	BEGIN_MESSAGE_MAP(CSortHeader, CHeaderCtrl)
	ON_MESSAGE(WM_MFCMAPI_SAVECOLUMNORDERHEADER, msgOnSaveColumnOrder)
#pragma warning(push)
#pragma warning( \
	disable : 26454) // Warning C26454 Arithmetic overflow: 'operator' operation produces a negative unsigned result at compile time
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnCustomDraw)
#pragma warning(pop)
	END_MESSAGE_MAP()

	_Check_return_ bool CSortHeader::Init(_In_ CHeaderCtrl* pHeader, _In_ HWND hwndParent)
	{
		if (!pHeader) return false;

		if (!SubclassWindow(pHeader->GetSafeHwnd()))
		{
			output::DebugPrint(output::dbgLevel::UI, L"CSortHeader::Init Unable to subclass existing header!\n");
			return false;
		}
		m_hwndParent = hwndParent;
		m_bTooltipDisplayed = false;

		RegisterHeaderTooltip();

		return true;
	}

	void CSortHeader::RegisterHeaderTooltip()
	{
		if (!m_hwndTip)
		{
			m_hwndTip = EC_D(
				HWND,
				CreateWindowEx(
					WS_EX_TOPMOST,
					TOOLTIPS_CLASS,
					nullptr,
					TTS_NOPREFIX | TTS_ALWAYSTIP,
					CW_USEDEFAULT,
					CW_USEDEFAULT,
					CW_USEDEFAULT,
					CW_USEDEFAULT,
					m_hWnd,
					nullptr,
					AfxGetInstanceHandle(),
					nullptr));

			if (m_hwndTip)
			{
				EC_B_S(::SetWindowPos(m_hwndTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE));

				m_ti.cbSize = sizeof(TOOLINFO);
				m_ti.uFlags = TTF_TRACK | TTF_IDISHWND;
				m_ti.hwnd = m_hWnd;
				m_ti.uId = reinterpret_cast<UINT_PTR>(m_hWnd);
				m_ti.hinst = AfxGetInstanceHandle();
				m_ti.lpszText = const_cast<LPWSTR>(L"");

				EC_B_S(::SendMessage(m_hwndTip, TTM_ADDTOOL, 0, LPARAM(&m_ti)));
				EC_B_S(::SendMessage(m_hwndTip, TTM_SETMAXTIPWIDTH, 0, LPARAM(500)));
			}
		}
	}

	LRESULT CSortHeader::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
	{
		if (m_hwndTip)
		{
			switch (message)
			{
			case WM_ERASEBKGND:
			{
				return true;
			}
			case WM_MOUSEMOVE:
				// This covers the case where we move from one header to another, but don't leave the control
				if (m_bTooltipDisplayed)
				{
					HDHITTESTINFO hdHitTestInfo = {0};
					hdHitTestInfo.pt.x = GET_X_LPARAM(lParam);
					hdHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

					WC_B_S(::SendMessage(m_hWnd, HDM_HITTEST, 0, reinterpret_cast<LPARAM>(&hdHitTestInfo)));
					if (!(hdHitTestInfo.flags & HHT_ONHEADER))
					{
						// We were displaying a tooltip, but we're now on empty space in the header. Turn it off.
						EC_B_S(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, false, reinterpret_cast<LPARAM>(&m_ti)));
						m_bTooltipDisplayed = false;
					}
				}
				else
				{
					TRACKMOUSEEVENT tmEvent = {0};
					tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
					tmEvent.dwFlags = TME_HOVER;
					tmEvent.hwndTrack = m_hWnd;
					tmEvent.dwHoverTime = HOVER_DEFAULT;

					EC_B_S(TrackMouseEvent(&tmEvent));
				}

				break;
			case WM_MOUSEHOVER:
			{
				HDHITTESTINFO hdHitTestInfo = {0};
				hdHitTestInfo.pt.x = GET_X_LPARAM(lParam);
				hdHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

				EC_B_S(::SendMessage(m_hWnd, HDM_HITTEST, 0, reinterpret_cast<LPARAM>(&hdHitTestInfo)));

				// We only turn on or modify our tooltip if we're on a column header
				if (hdHitTestInfo.flags & HHT_ONHEADER)
				{
					HDITEM hdItem = {0};
					hdItem.mask = HDI_LPARAM;

					EC_B_S(GetItem(hdHitTestInfo.iItem, &hdItem));

					const auto lpHeaderData = reinterpret_cast<LPHEADERDATA>(hdItem.lParam);

					// This will only display tips if we have a HeaderData structure saved
					if (lpHeaderData)
					{
						EC_B_S(::GetCursorPos(&hdHitTestInfo.pt));
						EC_B_S(::SendMessage(
							m_hwndTip,
							TTM_TRACKPOSITION,
							0,
							MAKELPARAM((DWORD_PTR) hdHitTestInfo.pt.x + 10, (DWORD_PTR) hdHitTestInfo.pt.y + 20)));

						m_ti.lpszText = const_cast<LPWSTR>(lpHeaderData->szTipString.c_str());
						EC_B_S(::SendMessage(m_hwndTip, TTM_SETTOOLINFOW, true, reinterpret_cast<LPARAM>(&m_ti)));
						// Ask for notification when the mouse leaves the control
						TRACKMOUSEEVENT tmEvent = {0};
						tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
						tmEvent.dwFlags = TME_LEAVE;
						tmEvent.hwndTrack = m_hWnd;
						EC_B_S(TrackMouseEvent(&tmEvent));

						// Turn on the tooltip
						EC_B_S(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, true, reinterpret_cast<LPARAM>(&m_ti)));
						m_bTooltipDisplayed = true;
					}
				}
				break;
			}
			case WM_MOUSELEAVE:
				// We were displaying a tooltip, but we're now off of the header. Turn it off.
				EC_B_S(::SendMessage(m_hwndTip, TTM_TRACKACTIVATE, false, reinterpret_cast<LPARAM>(&m_ti)));
				m_bTooltipDisplayed = false;
				return NULL;
			}
		}

		return CHeaderCtrl::WindowProc(message, wParam, lParam);
	}

	// WM_MFCMAPI_SAVECOLUMNORDERHEADER
	_Check_return_ LRESULT CSortHeader::msgOnSaveColumnOrder(WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		if (m_hwndParent)
		{
			WC_B_S(::PostMessage(
				m_hwndParent,
				WM_MFCMAPI_SAVECOLUMNORDERLIST,
				reinterpret_cast<WPARAM>(nullptr),
				reinterpret_cast<LPARAM>(nullptr)));
		}
		return S_OK;
	}

	void CSortHeader::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult) { ui::CustomDrawHeader(pNMHDR, pResult); }
} // namespace controls::sortlistctrl
