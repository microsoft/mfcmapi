#pragma once
#include <Dialogs/Dialog.h>
#include "ViewPane/ViewPane.h"
#include "ViewPane/TextPane.h"
#include "ViewPane/CheckPane.h"
#include "ViewPane/DropDownPane.h"
#include "ViewPane/ListPane.h"
#include "ViewPane/CollapsibleTextPane.h"

class CParentWnd;

// Buttons for CEditor
//#define CEDITOR_BUTTON_OK 0x00000001 // Duplicated from MFCMAPI.h - do not modify
#define CEDITOR_BUTTON_ACTION1 0x00000002
#define CEDITOR_BUTTON_ACTION2 0x00000004
//#define CEDITOR_BUTTON_CANCEL 0x00000008 // Duplicated from MFCMAPI.h - do not modify
#define CEDITOR_BUTTON_ACTION3 0x00000010

template<typename T>
DoListEditCallback ListEditCallBack(T* editor)
{
	return [editor](auto a, auto b, auto c) {return editor->DoListEdit(a, b, c); };
}

class CEditor : public CMyDialog
{
public:
	// Main Edit Constructor
	CEditor(
		_In_opt_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulButtonFlags);
	CEditor(
		_In_opt_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulButtonFlags,
		UINT uidActionButtonText1,
		UINT uidActionButtonText2,
		UINT uidActionButtonText3);
	virtual ~CEditor();

	_Check_return_ HRESULT DisplayDialog();

	// These functions can be used to set up a data editing dialog
	void SetPromptPostFix(_In_ wstring szMsg);
	void InitPane(ULONG iNum, ViewPane* lpPane);
	void SetStringA(ULONG i, string szMsg) const;
	void SetStringW(ULONG i, wstring szMsg) const;
	void SetStringf(ULONG i, LPCWSTR szMsg, ...) const;
#ifdef CHECKFORMATPARAMS
#undef SetStringf
#define SetStringf(i, fmt,...) SetStringf(i, fmt, __VA_ARGS__), wprintf(fmt, __VA_ARGS__)
#endif

	void LoadString(ULONG i, UINT uidMsg) const;
	void SetBinary(ULONG i, _In_opt_count_(cb) LPBYTE lpb, size_t cb) const;
	void SetBinary(ULONG i, _In_ vector<BYTE> bin) const;
	void SetBinary(ULONG i, _In_ SBinary bin) const;
	void SetHex(ULONG i, ULONG ulVal) const;
	void SetDecimal(ULONG i, ULONG ulVal) const;
	void SetSize(ULONG i, size_t cb) const;

	// Get values after we've done the DisplayDialog
	ViewPane* GetPane(ULONG iPane) const;
	wstring GetStringW(ULONG i) const;
	_Check_return_ ULONG GetHex(ULONG i) const;
	_Check_return_ ULONG GetDecimal(ULONG i) const;
	_Check_return_ ULONG GetPropTag(ULONG i) const;
	_Check_return_ bool GetCheck(ULONG i) const;
	_Check_return_ int GetDropDown(ULONG i) const;
	_Check_return_ DWORD_PTR GetDropDownValue(ULONG i) const;
	_Check_return_ HRESULT GetEntryID(ULONG i, bool bIsBase64, _Out_ size_t* lpcbBin, _Out_ LPENTRYID* lpEID) const;
	_Check_return_ GUID GetSelectedGUID(ULONG iControl, bool bByteSwapped) const;

	// AddIn functions
	void SetAddInTitle(wstring szTitle);
	void SetAddInLabel(ULONG i, wstring szLabel) const;

	// Use this function to implement list editing
	// return true to indicate the entry was changed, false to indicate it was not
	_Check_return_ virtual bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);

protected:
	// Functions used by derived classes during init
	void InsertColumn(ULONG ulListNum, int nCol, UINT uidText) const;
	void InsertColumn(ULONG ulListNum, int nCol, UINT uidText, ULONG ulPropType) const;
	void SetListString(ULONG iControl, ULONG iListRow, ULONG iListCol, wstring szListString) const;
	_Check_return_ SortListData* InsertListRow(ULONG iControl, int iRow, wstring szText) const;
	void ClearList(ULONG iControl) const;
	void ResizeList(ULONG uControl, bool bSort) const;

	// Functions used by derived classes during handle change events
	_Check_return_ ULONG GetHexUseControl(ULONG i) const;
	_Check_return_ ULONG GetDecimalUseControl(ULONG i) const;
	_Check_return_ ULONG GetPropTagUseControl(ULONG iControl) const;
	vector<BYTE> GetBinaryUseControl(ULONG i) const;
	_Check_return_ string GetEditBoxTextA(ULONG iControl) const;
	_Check_return_ wstring GetEditBoxTextW(ULONG iContro) const;
	_Check_return_ ULONG GetListCount(ULONG iControl) const;
	_Check_return_ SortListData* GetListRowData(ULONG iControl, int iRow) const;
	_Check_return_ bool IsDirty(ULONG iControl) const;

	// Called to enable/disable buttons based on number of items
	void UpdateButtons() const;
	BOOL OnInitDialog() override;
	void OnOK() override;
	void OnRecalcLayout();

	// protected since derived classes need to call the base implementation
	_Check_return_ virtual ULONG HandleChange(UINT nID);

	void EnableScroll();

private:
	// Overridable functions
	// use these functions to implement custom edit buttons
	virtual void OnEditAction1();
	virtual void OnEditAction2();
	virtual void OnEditAction3();

	// private constructor invoked from public constructors
	void Constructor(
		_In_opt_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulButtonFlags,
		UINT uidActionButtonText1,
		UINT uidActionButtonText2,
		UINT uidActionButtonText3);

	void DeleteControls();
	_Check_return_ SIZE ComputeWorkArea(SIZE sScreen);
	void OnGetMinMaxInfo(_Inout_ MINMAXINFO* lpMMI);
	void OnSetDefaultSize();
	_Check_return_ LRESULT OnNcHitTest(CPoint point);
	void OnSize(UINT nType, int cx, int cy);
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
	void SetMargins() const;

	// List functions and data
	_Check_return_ bool OnEditListEntry(ULONG ulListNum) const;
	ULONG m_ulListNum; // Only supporting one list right now - this is the control number for it

	// Our UI controls. Only valid during display.
	bool m_bHasPrompt;
	CEdit m_Prompt;
	CButton m_OkButton;
	CButton m_ActionButton1;
	CButton m_ActionButton2;
	CButton m_ActionButton3;
	CButton m_CancelButton;
	ULONG m_cButtons;

	// Variables that get set in the constructor
	ULONG m_bButtonFlags;

	// Size calculations
	int m_iMargin;
	int m_iSideMargin;
	int m_iSmallHeightMargin;
	int m_iLargeHeightMargin;
	int m_iButtonWidth;
	int m_iEditHeight;
	int m_iTextHeight;
	int m_iButtonHeight;
	int m_iMinWidth;
	int m_iMinHeight;

	// Title bar, prompt and icon
	UINT m_uidTitle;
	wstring m_szTitle;
	UINT m_uidPrompt;
	wstring m_szPromptPostFix;
	wstring m_szAddInTitle;
	HICON m_hIcon;

	UINT m_uidActionButtonText1;
	UINT m_uidActionButtonText2;
	UINT m_uidActionButtonText3;

	vector<ViewPane*> m_lpControls; // array of controls

	bool m_bEnableScroll;
	CWnd m_ScrollWindow;
	HWND m_hWndVertScroll;
	bool m_bScrollVisible;

	int m_iScrollClient;

	DECLARE_MESSAGE_MAP()
};