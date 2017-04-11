#pragma once
#include <UI/ViewPane/ViewPane.h>

class TextPane : public ViewPane
{
public:
	TextPane()
	{
		m_bMultiline = false;
		m_bCommitted = false;
	}

	static TextPane* CreateMultiLinePane(UINT uidLabel, bool bReadOnly);
	static TextPane* CreateMultiLinePane(UINT uidLabel, _In_ const wstring& szVal, bool bReadOnly);
	static TextPane* CreateSingleLinePane(UINT uidLabel, bool bReadOnly, bool bMultiLine = false);
	static TextPane* CreateSingleLinePane(UINT uidLabel, _In_ const wstring& szVal, bool bReadOnly, bool bMultiLine = false);
	static TextPane* CreateSingleLinePaneID(UINT uidLabel, UINT uidVal, bool bReadOnly);
	static TextPane* CreateCollapsibleTextPane(UINT uidLabel, bool bReadOnly);

	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	int GetFixedHeight() override;
	int GetLines() override;

	void Clear();
	void SetStringW(const wstring& szMsg);
	void SetBinary(_In_opt_count_(cb) LPBYTE lpb, size_t cb);
	void SetBinaryStream(_In_ LPSTREAM lpStreamIn);
	void GetBinaryStream(_In_ LPSTREAM lpStreamOut) const;
	void AppendString(_In_ const wstring& szMsg);
	void ShowWindow(int nCmdShow);

	void SetReadOnly();
	void SetMultiline();
	bool IsDirty() override;

	wstring GetStringW() const;

protected:
	CRichEditCtrl m_EditBox;
	static const int LINES_MULTILINEEDIT = 4;

private:
	wstring GetUIValue() const;
	void CommitUIValues() override;
	void SetEditBoxText();

	wstring m_lpszW;
	bool m_bCommitted;
	bool m_bMultiline;
};