#include <StdAfx.h>
#include <UI/Controls/ScrollHost.h>
#include <UI/UIFunctions.h>

namespace controls
{
	constexpr UINT_PTR scrollRepeatTimer = 0x4D46;
	constexpr UINT scrollRepeatInterval = 50;

	BOOL CScrollHost::Create(_In_ CWnd* parent)
	{
		if (!parent) return false;

		m_parent = parent->GetSafeHwnd();
		const auto created = CWnd::CreateEx(
			0,
			_T("Static"), // STRING_OK
			nullptr,
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_VSCROLL,
			CRect(0, 0, 0, 0),
			parent,
			NULL);
		if (!created) return false;

		ui::StyleScrollHost(m_hWnd);
		return ModifyStyleEx(0, WS_EX_CONTROLPARENT);
	}

	bool CScrollHost::GetScrollGeometry(_Out_ ScrollGeometry* geometry) const noexcept
	{
		if (!geometry) return false;

		auto info = SCROLLINFO{};
		info.cbSize = sizeof info;
		info.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
		if (!::GetScrollInfo(m_hWnd, SB_VERT, &info)) return false;

		auto window = RECT{};
		::GetWindowRect(m_hWnd, &window);
		const auto arrowHeight = GetSystemMetrics(SM_CYVSCROLL);
		geometry->trackTop = window.top + arrowHeight;
		geometry->trackHeight = max(0, window.bottom - window.top - 2 * arrowHeight);
		geometry->minimum = info.nMin;

		const auto range = max(1, info.nMax - info.nMin + 1);
		const auto page = min(static_cast<int>(info.nPage), range);
		geometry->maximumPosition = max(info.nMin, info.nMax - page + 1);
		geometry->thumbHeight = page ? MulDiv(geometry->trackHeight, page, range) : GetSystemMetrics(SM_CYVTHUMB);
		geometry->thumbHeight = min(geometry->trackHeight, max(GetSystemMetrics(SM_CYVTHUMB), geometry->thumbHeight));
		geometry->thumbTop = geometry->trackTop;
		if (geometry->maximumPosition > geometry->minimum && geometry->trackHeight > geometry->thumbHeight)
		{
			geometry->thumbTop += MulDiv(
				geometry->trackHeight - geometry->thumbHeight,
				info.nPos - geometry->minimum,
				geometry->maximumPosition - geometry->minimum);
		}

		return true;
	}

	void CScrollHost::StopTracking() noexcept
	{
		KillTimer(scrollRepeatTimer);
		m_dragOffset = -1;
		m_repeatCommand = SB_ENDSCROLL;
	}

	LRESULT CScrollHost::WindowProc(const UINT message, const WPARAM wParam, const LPARAM lParam)
	{
		LRESULT result = 0;
		if (ui::HandleControlUI(message, wParam, lParam, &result)) return result;

		switch (message)
		{
		case WM_NCHITTEST:
		{
			auto client = RECT{};
			GetClientRect(&client);
			auto clientRight = POINT{client.right, client.top};
			::ClientToScreen(m_hWnd, &clientRight);

			auto window = RECT{};
			GetWindowRect(&window);
			const auto point = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
			if (point.x >= clientRight.x && point.x < window.right && point.y >= window.top && point.y < window.bottom)
			{
				return HTVSCROLL;
			}

			break;
		}
		case WM_NCLBUTTONDOWN:
			if (wParam == HTVSCROLL)
			{
				auto geometry = ScrollGeometry{};
				if (!GetScrollGeometry(&geometry)) break;

				const auto y = GET_Y_LPARAM(lParam);
				auto window = RECT{};
				GetWindowRect(&window);
				auto command = SB_ENDSCROLL;
				if (y < geometry.trackTop)
					command = SB_LINEUP;
				else if (y >= window.bottom - GetSystemMetrics(SM_CYVSCROLL))
					command = SB_LINEDOWN;
				else if (y < geometry.thumbTop)
					command = SB_PAGEUP;
				else if (y >= geometry.thumbTop + geometry.thumbHeight)
					command = SB_PAGEDOWN;
				else
				{
					m_dragOffset = y - geometry.thumbTop;
					SetCapture();
				}

				if (command != SB_ENDSCROLL)
				{
					m_repeatCommand = command;
					SetCapture();
					SetTimer(scrollRepeatTimer, GetDoubleClickTime(), nullptr);
					return ::SendMessage(m_parent, WM_VSCROLL, static_cast<WPARAM>(command), 0);
				}

				return 0;
			}
			break;
		case WM_MOUSEMOVE:
			if (::GetCapture() == m_hWnd && m_dragOffset >= 0)
			{
				auto geometry = ScrollGeometry{};
				if (!GetScrollGeometry(&geometry)) return 0;

				auto point = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
				::ClientToScreen(m_hWnd, &point);
				const auto availableTrack = geometry.trackHeight - geometry.thumbHeight;
				const auto thumbTop = min(availableTrack, max(0, point.y - m_dragOffset - geometry.trackTop));
				m_trackPosition =
					availableTrack && geometry.maximumPosition > geometry.minimum
						? geometry.minimum +
							  MulDiv(thumbTop, geometry.maximumPosition - geometry.minimum, availableTrack)
						: geometry.minimum;
				return ::SendMessage(m_parent, WM_VSCROLL, static_cast<WPARAM>(SB_THUMBTRACK), 0);
			}
			break;
		case WM_TIMER:
			if (wParam == scrollRepeatTimer)
			{
				auto geometry = ScrollGeometry{};
				auto window = RECT{};
				auto point = POINT{};
				if (!GetScrollGeometry(&geometry) || !::GetWindowRect(m_hWnd, &window) || !::GetCursorPos(&point))
					return 0;

				const auto arrowHeight = GetSystemMetrics(SM_CYVSCROLL);
				const auto pointerInCommand =
					(m_repeatCommand == SB_LINEUP && point.y < geometry.trackTop) ||
					(m_repeatCommand == SB_LINEDOWN && point.y >= window.bottom - arrowHeight) ||
					(m_repeatCommand == SB_PAGEUP && point.y >= geometry.trackTop && point.y < geometry.thumbTop) ||
					(m_repeatCommand == SB_PAGEDOWN && point.y >= geometry.thumbTop + geometry.thumbHeight &&
					 point.y < window.bottom - arrowHeight);
				if (pointerInCommand)
				{
					::SendMessage(m_parent, WM_VSCROLL, static_cast<WPARAM>(m_repeatCommand), 0);
				}

				KillTimer(scrollRepeatTimer);
				SetTimer(scrollRepeatTimer, scrollRepeatInterval, nullptr);
				return 0;
			}
			break;
		case WM_LBUTTONUP:
			if (::GetCapture() == m_hWnd)
			{
				ReleaseCapture();
				StopTracking();
				return 0;
			}
			break;
		case WM_CAPTURECHANGED:
			StopTracking();
			break;
		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case BN_SETFOCUS:
			case EN_SETFOCUS:
				return ::SendMessage(m_parent, message, wParam, lParam);
			}
			break;
		case WM_NEXTDLGCTL:
		case WM_MOUSEWHEEL:
			return ::SendMessage(m_parent, message, wParam, lParam);
		case WM_VSCROLL:
			return ::SendMessage(m_parent, message, wParam, 0);
		case WM_NCDESTROY:
			StopTracking();
			break;
		}

		return CWnd::WindowProc(message, wParam, lParam);
	}
} // namespace controls
