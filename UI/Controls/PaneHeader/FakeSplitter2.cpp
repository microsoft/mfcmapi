#include <StdAfx.h>
#include <UI/Controls/PaneHeader/FakeSplitter2.h>
#include <UI/UIFunctions.h>
#include <UI/DoubleBuffer.h>
#include <core/utility/output.h>

namespace controls
{
	static std::wstring CLASS = L"CFakeSplitter2";

	enum FakesSplitHitTestValue
	{
		noHit = 0,
		SplitterHit = 1
	};

	CFakeSplitter2::~CFakeSplitter2()
	{
		TRACE_DESTRUCTOR(CLASS);
		CWnd::DestroyWindow();
	}

	void CFakeSplitter2::Init(HWND hWnd)
	{
		m_hwndParent = hWnd;
		WNDCLASSEX wc = {};
		const auto hInst = AfxGetInstanceHandle();
		if (!::GetClassInfoEx(hInst, _T("FakeSplitter"), &wc)) // STRING_OK
		{
			wc.cbSize = sizeof wc;
			wc.style = 0; // not passing CS_VREDRAW | CS_HREDRAW fixes flicker
			wc.lpszClassName = _T("FakeSplitter"); // STRING_OK
			wc.lpfnWndProc = ::DefWindowProc;
			wc.hbrBackground = GetSysBrush(ui::uiColor::Background); // helps spot flashing

			RegisterClassEx(&wc);
		}

		// WS_CLIPCHILDREN is used to reduce flicker
		EC_B_S(CreateEx(
			0,
			_T("FakeSplitter"), // STRING_OK
			_T("FakeSplitter"), // STRING_OK
			WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_VISIBLE,
			0,
			0,
			0,
			0,
			m_hwndParent,
			reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_FAKE_SPLITTER2)),
			nullptr));

		// Necessary for TAB to work. Without this, all TABS get stuck on the fake splitter control
		// instead of passing to the children. Haven't tested with nested splitters.
		EC_B_S(ModifyStyleEx(0, WS_EX_CONTROLPARENT));
	}

	BEGIN_MESSAGE_MAP(CFakeSplitter2, CWnd)
	ON_WM_SIZE()
	ON_WM_PAINT()
	END_MESSAGE_MAP()

	LRESULT CFakeSplitter2::WindowProc(const UINT message, const WPARAM wParam, const LPARAM lParam)
	{
		LRESULT lRes = 0;
		if (ui::HandleControlUI(message, wParam, lParam, &lRes)) return lRes;

		switch (message)
		{
		case WM_CLOSE:
			::SendMessage(m_hwndParent, message, wParam, lParam);
			return true;
		case WM_HELP:
			return true;
		case WM_COMMAND:
		{
			const auto nCode = HIWORD(wParam);
			if (EN_CHANGE == nCode || CBN_SELCHANGE == nCode || CBN_EDITCHANGE == nCode || BN_CLICKED == nCode)
			{
				::SendMessage(m_hwndParent, message, wParam, lParam);
			}

			break;
		}
		}

		return CWnd::WindowProc(message, wParam, lParam);
	}

	void CFakeSplitter2::SetPaneOne(HWND paneOne) noexcept { m_PaneOne = paneOne; }

	void CFakeSplitter2::SetPaneTwo(HWND /*paneTwo*/) noexcept {}

	void CFakeSplitter2::OnSize(UINT /*nType*/, const int cx, const int cy)
	{
		auto hdwp = WC_D(HDWP, BeginDeferWindowPos(2));
		if (hdwp)
		{
			hdwp = EC_D(HDWP, DeferWindowPos(hdwp, 0, 0, cx, cy));
			EC_B_S(EndDeferWindowPos(hdwp));
		}
	}

	HDWP CFakeSplitter2::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		output::DebugPrint(
			output::dbgLevel::Draw,
			L"CFakeSplitter2::DeferWindowPos x:%d y:%d width:%d height:%d v:%d\n",
			x,
			y,
			width,
			height,
			IsWindowVisible());
		InvalidateRect(CRect(x, y, width, height), false);
		if (m_PaneOne || m_ViewPaneOne)
		{
			CRect r1;
			r1.SetRect(x, y, width, height);

			if (m_PaneOne)
			{
				hWinPosInfo = ui::DeferWindowPos(
					hWinPosInfo, m_PaneOne, x, y, r1.Width(), r1.Height(), L"CFakeSplitter2::DeferWindowPos pane 1");
			}

			if (m_ViewPaneOne)
			{
				output::DebugPrint(
					output::dbgLevel::Draw,
					L"CFakeSplitter2::DeferWindowPos viewpane 1 x:%d y:%d width:%d height:%d\n",
					x,
					y,
					width,
					height);
				hWinPosInfo = EC_D(HDWP, m_ViewPaneOne->DeferWindowPos(hWinPosInfo, x, y, r1.Width(), r1.Height()));
			}
		}

		output::DebugPrint(output::dbgLevel::Draw, L"CFakeSplitter2::DeferWindowPos end\n");
		return hWinPosInfo;
	}

	void CFakeSplitter2::OnPaint()
	{
		auto ps = PAINTSTRUCT{};
		::BeginPaint(m_hWnd, &ps);
		if (ps.hdc)
		{
			auto rcWin = RECT{};
			::GetClientRect(m_hWnd, &rcWin);
			ui::CDoubleBuffer db;
			auto hdc = ps.hdc;
			db.Begin(hdc, rcWin);
			FillRect(hdc, &rcWin, GetSysBrush(ui::uiColor::Background));
			db.End(hdc);
		}

		::EndPaint(m_hWnd, &ps);
	}
} // namespace controls