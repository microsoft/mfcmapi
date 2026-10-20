#pragma once
#include <UI/ViewPane/TextPane.h>

namespace viewpane
{
	enum class pathType
	{
		directory,
		fileOpen,
		fileSave
	};

	class PathPane final : public TextPane
	{
	public:
		static std::shared_ptr<PathPane> Create(
			int paneID,
			UINT uidLabel,
			pathType type,
			const std::wstring& path = {},
			const std::wstring& extension = {},
			const std::wstring& filter = {},
			DWORD flags = 0);
		void SetHelpText(UINT uidHelp) { m_helpText = strings::loadstring(uidHelp); }
		int GetMinWidth() override;
		int GetFixedHeight() override;
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		bool containsWindow(HWND hWnd) const noexcept override;
		RECT GetWindowRect() const noexcept override;
		ULONG HandleChange(UINT nID) override;

	private:
		void Initialize(_In_ CWnd* parent, _In_ HDC hdc) override;
		int GetHelpHeight() const noexcept;
		std::wstring BrowseDirectory();
		static int CALLBACK BrowseCallback(HWND window, UINT message, LPARAM parameter, LPARAM data);

		pathType m_type{};
		UINT m_uidLabel{};
		UINT m_browseID{};
		std::wstring m_extension;
		std::wstring m_filter;
		DWORD m_flags{};
		std::wstring m_helpText;
		CStatic m_help;
		SIZE m_helpSize{};
	};
} // namespace viewpane
