#pragma once
#include <UI/ViewPane/ViewPane.h>

namespace viewpane
{
	class TextPane : public ViewPane
	{
	public:
		static TextPane* CreateMultiLinePane(UINT uidLabel, bool bReadOnly);
		static TextPane* CreateMultiLinePane(UINT uidLabel, _In_ const std::wstring& szVal, bool bReadOnly);
		static TextPane* CreateSingleLinePane(UINT uidLabel, bool bReadOnly, bool bMultiLine = false);
		static TextPane*
		CreateSingleLinePane(UINT uidLabel, _In_ const std::wstring& szVal, bool bReadOnly, bool bMultiLine = false);
		static TextPane* CreateSingleLinePaneID(UINT uidLabel, UINT uidVal, bool bReadOnly);
		static TextPane* CreateCollapsibleTextPane(UINT uidLabel, bool bReadOnly);

		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetFixedHeight() override;
		int GetLines() override;

		void Clear();
		void SetStringW(const std::wstring& szMsg);
		void SetBinary(_In_opt_count_(cb) LPBYTE lpb, size_t cb);
		void SetBinaryStream(_In_ LPSTREAM lpStreamIn);
		void GetBinaryStream(_In_ LPSTREAM lpStreamOut) const;
		void AppendString(_In_ const std::wstring& szMsg);
		void ShowWindow(int nCmdShow);

		void SetReadOnly();
		void SetMultiline();
		bool IsDirty() override;

		std::wstring GetStringW() const;

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
	};
} // namespace viewpane