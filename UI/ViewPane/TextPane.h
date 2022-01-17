#pragma once
#include <UI/ViewPane/ViewPane.h>

namespace viewpane
{
	struct Range
	{
		size_t start;
		size_t end;
	};

	class TextPane : public ViewPane
	{
	public:
		static std::shared_ptr<TextPane> CreateMultiLinePane(int paneID, UINT uidLabel, bool bReadOnly);
		static std::shared_ptr<TextPane>
		CreateMultiLinePane(int paneID, UINT uidLabel, _In_ const std::wstring& szVal, bool bReadOnly);
		static std::shared_ptr<TextPane>
		CreateSingleLinePane(int paneID, UINT uidLabel, bool bReadOnly, bool bMultiLine = false);
		static std::shared_ptr<TextPane> CreateSingleLinePane(
			const int paneID,
			_In_ const std::wstring& szLabel,
			_In_ const std::wstring& szVal,
			const bool bReadOnly,
			const bool bMultiLine = false);
		static std::shared_ptr<TextPane> CreateSingleLinePane(
			int paneID,
			UINT uidLabel,
			_In_ const std::wstring& szVal,
			bool bReadOnly,
			bool bMultiLine = false);
		static std::shared_ptr<TextPane> CreateSingleLinePaneID(int paneID, UINT uidLabel, UINT uidVal, bool bReadOnly);
		static std::shared_ptr<TextPane> CreateCollapsibleTextPane(int paneID, UINT uidLabel, bool bReadOnly);

		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetFixedHeight() override;
		int GetLines() override;

		void Clear();
		void SetStringW(const std::wstring& szMsg);
		void SetBinary(_In_opt_count_(cb) const BYTE* lpb, size_t cb);
		void SetBinaryStream(_In_ LPSTREAM lpStreamIn);
		void GetBinaryStream(_In_ LPSTREAM lpStreamOut) const;
		void AppendString(_In_ const std::wstring& szMsg);
		void ShowWindow(const int nCmdShow) { m_EditBox.ShowWindow(nCmdShow); }

		void SetReadOnly();
		void SetMultiline();
		bool IsDirty() override;

		std::wstring GetStringW() const;
		void AddHighlight(const Range& range)
		{
			m_highlights.emplace_back(range);
			DoHighlights();
		}
		void ClearHighlight()
		{
			m_highlights.clear();
			DoHighlights();
		}
		void DoHighlights();
		bool containsWindow(HWND hWnd) const noexcept override;
		RECT GetWindowRect() const noexcept override
		{
			auto rcEdit = RECT{};
			::GetWindowRect(m_EditBox.GetSafeHwnd(), &rcEdit);
			const auto rcHeader = ViewPane::GetWindowRect();
			auto rcPane = RECT{};
			::UnionRect(&rcPane, &rcEdit, &rcHeader);
			return rcPane;
		}

	protected:
		CRichEditCtrl m_EditBox;
		static const int LINES_MULTILINEEDIT = 4;

	private:
		std::wstring GetUIValue() const;
		void CommitUIValues() override;
		void SetEditBoxText();

		std::wstring m_lpszW;
		bool m_bCommitted{false};
		bool m_bMultiline{false};
		std::vector<Range> m_highlights;
	};
} // namespace viewpane