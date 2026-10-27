#include <StdAfx.h>
#include <UI/ViewPane/ThemePreviewPane.h>
#include <UI/UIFunctions.h>

namespace viewpane
{
	std::shared_ptr<ThemePreviewPane> ThemePreviewPane::Create(const int paneID, const UINT uidLabel)
	{
		auto pane = std::make_shared<ThemePreviewPane>();
		if (pane)
		{
			pane->m_paneID = paneID;
			pane->SetLabel(uidLabel);
		}

		return pane;
	}

	void
	ThemePreviewPane::SetTheme(const ui::theme::ThemeDefinition& definition, const ui::theme::Appearance appearance)
	{
		m_themeName = definition.name;
		const auto effectiveAppearance = ui::GetEffectiveTheme() == ui::uiTheme::HighContrast
											 ? ui::uiTheme::HighContrast
										 : appearance == ui::theme::Appearance::Dark ? ui::uiTheme::Dark
																					 : ui::uiTheme::Light;
		for (size_t index = 0; index < m_colors.size(); ++index)
		{
			m_colors[index] = ui::theme::GetColor(definition, effectiveAppearance, static_cast<ui::uiColor>(index));
		}

		if (m_hWndPreview) InvalidateRect(m_hWndPreview, nullptr, true);
	}

	void ThemePreviewPane::Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);

		m_hWndPreview = EC_D(
			HWND,
			::CreateWindowExW(
				0,
				L"STATIC",
				L"",
				WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
				0,
				0,
				0,
				0,
				pParent->GetSafeHwnd(),
				reinterpret_cast<HMENU>(static_cast<INT_PTR>(m_nID)),
				AfxGetInstanceHandle(),
				nullptr));

		EC_B_S(::SetWindowSubclass(m_hWndPreview, PreviewProc, 0, reinterpret_cast<DWORD_PTR>(this)));
		m_bInitialized = true;
	}

	HDWP
	ThemePreviewPane::DeferWindowPos(_In_ HDWP hWinPosInfo, const int x, const int y, const int width, const int height)
	{
		hWinPosInfo = ViewPane::DeferWindowPos(hWinPosInfo, x, y, width, height);

		return ui::DeferWindowPos(
			hWinPosInfo,
			m_hWndPreview,
			x,
			y + GetHeaderHeight(),
			width,
			height - GetHeaderHeight(),
			L"ThemePreviewPane::DeferWindowPos");
	}

	int ThemePreviewPane::GetMinWidth() { return max(ViewPane::GetMinWidth(), 18 * m_iEditHeight); }

	int ThemePreviewPane::GetFixedHeight() { return GetHeaderHeight() + 12 * m_iEditHeight + 6 * max(6, m_iMargin); }

	bool ThemePreviewPane::containsWindow(const HWND hWnd) const noexcept
	{
		return m_hWndPreview == hWnd || ViewPane::containsWindow(hWnd);
	}

	RECT ThemePreviewPane::GetWindowRect() const noexcept
	{
		auto rcPreview = RECT{};
		::GetWindowRect(m_hWndPreview, &rcPreview);
		const auto rcHeader = ViewPane::GetWindowRect();
		auto rcPane = RECT{};
		::UnionRect(&rcPane, &rcPreview, &rcHeader);

		return rcPane;
	}

	LRESULT CALLBACK ThemePreviewPane::PreviewProc(
		const HWND hWnd,
		const UINT message,
		const WPARAM wParam,
		const LPARAM lParam,
		const UINT_PTR subclassID,
		const DWORD_PTR data)
	{
		const auto pane = reinterpret_cast<ThemePreviewPane*>(data);

		if (message == WM_NCDESTROY)
		{
			RemoveWindowSubclass(hWnd, PreviewProc, subclassID);
			pane->m_hWndPreview = nullptr;

			return DefSubclassProc(hWnd, message, wParam, lParam);
		}

		if (message == WM_ERASEBKGND) return TRUE;

		if (message == WM_SIZE)
		{
			const auto result = DefSubclassProc(hWnd, message, wParam, lParam);
			InvalidateRect(hWnd, nullptr, TRUE);

			return result;
		}

		if (message != WM_PAINT) return DefSubclassProc(hWnd, message, wParam, lParam);

		auto paint = PAINTSTRUCT{};
		const auto hdc = BeginPaint(hWnd, &paint);

		auto client = RECT{};
		GetClientRect(hWnd, &client);

		const auto fillRect = [hdc](const RECT& rect, const COLORREF color) {
			const auto brush = CreateSolidBrush(color);
			FillRect(hdc, &rect, brush);
			DeleteObject(brush);
		};

		const auto drawFrame = [hdc](const RECT& rect, const COLORREF color, const int width) {
			const auto brush = CreateSolidBrush(color);
			auto inset = rect;

			for (auto i = 0; i < width; ++i)
			{
				FrameRect(hdc, &inset, brush);
				InflateRect(&inset, -1, -1);
			}

			DeleteObject(brush);
		};

		const auto getColor = [pane](const ui::uiColor role) { return pane->m_colors[static_cast<size_t>(role)]; };
		const auto margin = max(6, pane->m_iMargin);
		const auto rowHeight = 2 * pane->m_iEditHeight;
		const auto oldFont = SelectObject(hdc, ui::GetUIFont());
		SetBkMode(hdc, TRANSPARENT);

		const auto drawText = [&](RECT rect, const std::wstring& label, const ui::uiColor foreground) {
			InflateRect(&rect, -margin, 0);
			SetTextColor(hdc, getColor(foreground));
			DrawTextW(hdc, label.c_str(), -1, &rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);
		};

		const auto drawSample = [&](const RECT& rect,
									const UINT label,
									const ui::uiColor background,
									const ui::uiColor foreground,
									const ui::uiColor border) {
			fillRect(rect, getColor(background));
			drawFrame(rect, getColor(border), 1);
			drawText(rect, strings::loadstring(label), foreground);
		};

		fillRect(client, getColor(ui::uiColor::Background));

		auto header = RECT{margin, margin, client.right - margin, margin + rowHeight};
		fillRect(header, getColor(ui::uiColor::PaneHeaderBackground));
		drawFrame(header, getColor(ui::uiColor::Glow), 2);
		drawText(header, pane->m_themeName, ui::uiColor::PaneHeaderText);

		const auto columnWidth = max(0, (client.right - 3 * margin) / 2);

		for (auto row = 0; row < 4; ++row)
		{
			const auto top = header.bottom + margin + row * (rowHeight + margin);
			const auto left = RECT{margin, top, margin + columnWidth, top + rowHeight};
			const auto right = RECT{left.right + margin, top, client.right - margin, top + rowHeight};

			switch (row)
			{
			case 0:
				drawSample(
					left,
					IDS_THEME_PREVIEW_NORMAL,
					ui::uiColor::Background,
					ui::uiColor::Text,
					ui::uiColor::FrameUnselected);
				drawSample(
					right,
					IDS_THEME_PREVIEW_READONLY,
					ui::uiColor::BackgroundReadOnly,
					ui::uiColor::TextReadOnly,
					ui::uiColor::FrameUnselected);
				break;
			case 1:
				drawSample(
					left,
					IDS_THEME_PREVIEW_DISABLED,
					ui::uiColor::BackgroundReadOnly,
					ui::uiColor::TextDisabled,
					ui::uiColor::FrameUnselected);
				drawSample(
					right,
					IDS_THEME_PREVIEW_FOCUS,
					ui::uiColor::Background,
					ui::uiColor::Text,
					ui::uiColor::FrameSelected);
				break;
			case 2:
				drawSample(
					left,
					IDS_THEME_PREVIEW_HOVER,
					ui::uiColor::GlowBackground,
					ui::uiColor::GlowText,
					ui::uiColor::Glow);
				drawSample(
					right,
					IDS_THEME_PREVIEW_SELECTION,
					ui::uiColor::SelectedBackground,
					ui::uiColor::GlowText,
					ui::uiColor::FrameSelected);
				break;
			case 3:
			{
				drawSample(
					left,
					IDS_THEME_PREVIEW_HIGHLIGHT,
					ui::uiColor::TextHighlightBackground,
					ui::uiColor::TextHighlight,
					ui::uiColor::FrameUnselected);
				fillRect(right, getColor(ui::uiColor::Background));
				auto gutter = right;
				gutter.right = min(right.right, right.left + rowHeight);
				fillRect(gutter, getColor(ui::uiColor::MenuGutter));
				drawText(gutter, L">", ui::uiColor::Arrow);

				auto label = right;
				label.left = gutter.right;
				drawText(label, strings::loadstring(IDS_THEME_PREVIEW_MENU), ui::uiColor::Text);
				drawFrame(right, getColor(ui::uiColor::FrameUnselected), 1);
				break;
			}
			}
		}

		auto footer = RECT{
			margin,
			header.bottom + 5 * margin + 4 * rowHeight,
			client.right - margin,
			header.bottom + 5 * margin + 5 * rowHeight};
		fillRect(footer, getColor(ui::uiColor::Status));
		drawText(footer, L"MFCMAPI", ui::uiColor::StatusText);
		SelectObject(hdc, oldFont);

		EndPaint(hWnd, &paint);

		return 0;
	}
} // namespace viewpane
