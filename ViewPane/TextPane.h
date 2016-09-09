#pragma once
// TextPane.h : header file

#include "ViewPane.h"

ViewPane* CreateMultiLinePane(UINT uidLabel, bool bReadOnly);
ViewPane* CreateMultiLinePane(UINT uidLabel, _In_ wstring szVal, bool bReadOnly);
ViewPane* CreateSingleLinePane(UINT uidLabel, bool bReadOnly, bool bMultiLine = false);
ViewPane* CreateSingleLinePane(UINT uidLabel, _In_ wstring szVal, bool bReadOnly, bool bMultiLine = false);
ViewPane* CreateSingleLinePaneID(UINT uidLabel, UINT uidVal, bool bReadOnly);

#define LINES_MULTILINEEDIT 4

class TextPane : public ViewPane
{
public:
	TextPane(UINT uidLabel, bool bReadOnly, bool bMultiLine);
	virtual ~TextPane();

	bool IsType(__ViewTypes vType) override;
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	void CommitUIValues() override;
	ULONG GetFlags() override;
	int GetFixedHeight() override;
	int GetLines() override;

	void ClearView();
	virtual void SetStringA(_In_opt_z_ LPCSTR szMsg, size_t cchsz = -1);
	virtual void SetStringW(_In_opt_z_ LPCWSTR szMsg, size_t cchsz = -1);
	void SetBinary(_In_opt_count_(cb) LPBYTE lpb, size_t cb);
	void InitEditFromBinaryStream(_In_ LPSTREAM lpStreamIn);
	void WriteToBinaryStream(_In_ LPSTREAM lpStreamOut) const;
	void AppendString(_In_z_ LPCTSTR szMsg);
	void ShowWindow(int nCmdShow);

	void SetEditReadOnly();

	wstring GetStringW() const;
	LPSTR GetStringA();
	_Check_return_ LPSTR GetEditBoxTextA(_Out_ size_t* lpcchText);
	_Check_return_ wstring GetEditBoxTextW(_Out_ size_t* lpcchText);
	_Check_return_ wstring GetStringUseControl() const;

protected:
	CRichEditCtrl m_EditBox;

private:
	void SetEditBoxText();
	void ClearString();

	LPWSTR m_lpszW;
	LPSTR m_lpszA; // on demand conversion of lpszW
	size_t m_cchsz; // length of string - maintained to preserve possible internal NULLs, includes NULL terminator
	bool m_bMultiline;
};