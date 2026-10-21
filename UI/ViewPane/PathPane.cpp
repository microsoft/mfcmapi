#include <StdAfx.h>
#include <UI/ViewPane/PathPane.h>
#include <UI/file/FileDialogEx.h>
#include <UI/UIFunctions.h>

namespace viewpane
{
	std::shared_ptr<PathPane> PathPane::Create(
		const int paneID,
		const UINT uidLabel,
		const pathType type,
		const std::wstring& path,
		const std::wstring& extension,
		const std::wstring& filter,
		const DWORD flags)
	{
		auto pane = std::make_shared<PathPane>();
		pane->m_paneID = paneID;
		pane->SetLabel(uidLabel);
		pane->ViewPane::SetReadOnly(false);
		pane->SetStringW(path);
		pane->m_type = type;
		pane->m_uidLabel = uidLabel;
		pane->m_extension = extension;
		pane->m_filter = filter;
		pane->m_flags = flags;
		return pane;
	}

	void PathPane::Initialize(_In_ CWnd* parent, _In_ HDC hdc)
	{
		// Use the reserved header ID, distinct from this pane's edit control and other browse buttons.
		m_browseID = LOWORD(IDC_PROP_CONTROL_ID_BASE + 2 * m_paneID);
		m_Header.EnableActionButton(m_browseID);
		TextPane::Initialize(parent, hdc);
		m_Header.SetActionButton(strings::loadstring(IDS_BROWSE));
		if (!m_helpText.empty())
		{
			EC_B_S(m_help.Create(
				nullptr,
				WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX,
				CRect(0, 0, 0, 0),
				parent,
				static_cast<UINT>(IDC_STATIC)));
			::SetWindowTextW(m_help.GetSafeHwnd(), m_helpText.c_str());
			ui::SubclassLabel(m_help.GetSafeHwnd());
			auto rect = RECT{};
			WC_D_S(::DrawTextW(hdc, m_helpText.c_str(), -1, &rect, DT_CALCRECT | DT_NOPREFIX));
			m_helpSize = SIZE{rect.right - rect.left, rect.bottom - rect.top};
		}
	}

	int PathPane::GetMinWidth() { return max(TextPane::GetMinWidth(), m_helpSize.cx); }

	int PathPane::GetHelpHeight() const noexcept
	{
		return m_helpSize.cy && !collapsed() ? m_iSmallHeightMargin + m_helpSize.cy : 0;
	}

	int PathPane::GetFixedHeight() { return TextPane::GetFixedHeight() + GetHelpHeight(); }

	HDWP PathPane::DeferWindowPos(_In_ HDWP hWinPosInfo, const int x, const int y, const int width, const int height)
	{
		hWinPosInfo = TextPane::DeferWindowPos(hWinPosInfo, x, y, width, height - GetHelpHeight());
		if (m_help.GetSafeHwnd())
		{
			m_help.ShowWindow(collapsed() ? SW_HIDE : SW_SHOW);
			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo,
				m_help.GetSafeHwnd(),
				x,
				y + height - m_helpSize.cy,
				width,
				m_helpSize.cy,
				L"PathPane::DeferWindowPos::help");
		}

		return hWinPosInfo;
	}

	bool PathPane::containsWindow(HWND hWnd) const noexcept
	{
		return (m_help.GetSafeHwnd() && m_help.GetSafeHwnd() == hWnd) || TextPane::containsWindow(hWnd);
	}

	RECT PathPane::GetWindowRect() const noexcept
	{
		auto rect = TextPane::GetWindowRect();
		if (m_help.GetSafeHwnd() && !collapsed())
		{
			auto help = RECT{};
			::GetWindowRect(m_help.GetSafeHwnd(), &help);
			::UnionRect(&rect, &rect, &help);
		}

		return rect;
	}

	int CALLBACK PathPane::BrowseCallback(HWND window, const UINT message, LPARAM /*parameter*/, const LPARAM data)
	{
		if (message == BFFM_INITIALIZED && data) ::SendMessageW(window, BFFM_SETSELECTIONW, TRUE, data);
		return 0;
	}

	std::wstring PathPane::BrowseDirectory()
	{
		const auto result = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
		if (FAILED(result))
		{
			CHECKHRES(result);
			return {};
		}

		const auto title = strings::loadstring(m_uidLabel);
		const auto current = GetStringW();
		BROWSEINFOW browse{};
		browse.hwndOwner = ::GetAncestor(m_EditBox.GetSafeHwnd(), GA_ROOT);
		browse.lpszTitle = title.c_str();
		browse.ulFlags = BIF_USENEWUI | BIF_RETURNONLYFSDIRS;
		browse.lpfn = BrowseCallback;
		browse.lParam = reinterpret_cast<LPARAM>(current.c_str());
		const auto item = SHBrowseForFolderW(&browse);
		std::wstring selected;
		if (item)
		{
			WCHAR directory[MAX_PATH]{};
			if (SUCCEEDED(EC_B(SHGetPathFromIDListW(item, directory)))) selected = directory;
			CoTaskMemFree(item);
		}

		CoUninitialize();
		return selected;
	}

	ULONG PathPane::HandleChange(const UINT nID)
	{
		if (nID != m_browseID) return TextPane::HandleChange(nID);

		std::wstring selected;
		const auto flags = m_flags | OFN_HIDEREADONLY | OFN_NOCHANGEDIR;
		const auto filter = m_filter.empty() ? strings::loadstring(IDS_ALLFILES) : m_filter;
		auto* parent = CWnd::FromHandle(::GetAncestor(m_EditBox.GetSafeHwnd(), GA_ROOT));
		switch (m_type)
		{
		case pathType::directory:
			selected = BrowseDirectory();
			break;
		case pathType::fileOpen:
			selected = file::CFileDialogExW::OpenFile(m_extension, GetStringW(), flags, filter, parent);
			break;
		case pathType::fileSave:
			selected =
				file::CFileDialogExW::SaveAs(m_extension, GetStringW(), flags | OFN_OVERWRITEPROMPT, filter, parent);
			break;
		}

		if (!selected.empty())
		{
			SetStringW(selected);
			m_EditBox.SetModify(true);
		}

		return GetID();
	}
} // namespace viewpane
