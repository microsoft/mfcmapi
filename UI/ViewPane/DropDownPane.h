#pragma once
#include <UI/ViewPane/ViewPane.h>

class DropDownPane : public ViewPane
{
public:
	static DropDownPane* Create(UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, bool bReadOnly);
	static DropDownPane* CreateGuid(UINT uidLabel, bool bReadOnly);

	void SetDropDownSelection(_In_ const std::wstring& szText);
	void InsertDropString(_In_ const std::wstring& szText, ULONG ulValue);
	_Check_return_ std::wstring GetDropStringUseControl() const;
	_Check_return_ int GetDropDownSelection() const;
	_Check_return_ DWORD_PTR GetDropDownSelectionValue() const;
	GUID GetSelectedGUID(bool bByteSwapped) const;
	_Check_return_ int GetDropDown() const;
	_Check_return_ DWORD_PTR GetDropDownValue() const;

protected:
	DropDownPane();
	void SetSelection(DWORD_PTR iSelection);
	void CreateControl(int iControl, _In_ CWnd* pParent, _In_ HDC hdc);
	void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;

	CComboBox m_DropDown;

private:
	void SetWindowPos(int x, int y, int width, int height) override;
	void CommitUIValues() override;
	int GetMinWidth(_In_ HDC hdc) override;
	int GetFixedHeight() override;

	vector<std::pair<std::wstring, ULONG>> m_DropList;
	std::wstring m_lpszSelectionString;
	int m_iDropSelection;
	DWORD_PTR m_iDropSelectionValue;
};