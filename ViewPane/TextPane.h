#pragma once
// TextPane.h : header file

#include "ViewPane.h"

#define LINES_MULTILINEEDIT 4

class TextPane : public ViewPane
{
public:
	static TextPane* CreateMultiLinePane(UINT uidLabel, bool bReadOnly);
	static TextPane* CreateMultiLinePane(UINT uidLabel, _In_ wstring szVal, bool bReadOnly);
	static TextPane* CreateSingleLinePane(UINT uidLabel, bool bReadOnly, bool bMultiLine = false);
	static TextPane* CreateSingleLinePane(UINT uidLabel, _In_ wstring szVal, bool bReadOnly, bool bMultiLine = false);
	static TextPane* CreateSingleLinePaneID(UINT uidLabel, UINT uidVal, bool bReadOnly);

	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
	void SetWindowPos(int x, int y, int width, int height) override;
	ULONG GetFlags() override;
	int GetFixedHeight() override;
	int GetLines() override;
	void ClearView();
	virtual void SetStringA(string szMsg);
	virtual void SetStringW(wstring szMsg);
	void SetBinary(_In_opt_count_(cb) LPBYTE lpb, size_t cb);
	void InitEditFromBinaryStream(_In_ LPSTREAM lpStreamIn);
	void WriteToBinaryStream(_In_ LPSTREAM lpStreamOut);
	void AppendString(_In_ wstring szMsg);
	void ShowWindow(int nCmdShow);

	void SetEditReadOnly();

	wstring GetStringW() const;
	_Check_return_ string GetEditBoxTextA();
	_Check_return_ wstring GetEditBoxTextW();
	_Check_return_ wstring GetStringUseControl();
	bool m_bMultiline;

protected:
	bool IsType(__ViewTypes vType) override;
	CRichEditCtrl m_EditBox;

private:
	void CommitUIValues() override;
	string GetStringA() const;
	void SetEditBoxText();

	wstring m_lpszW;
};