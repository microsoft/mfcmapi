#pragma once
#include "Dialog.h"
#include "enums.h"
#include "ViewPane\ViewPane.h"
#include "ViewPane\TextPane.h"
#include "ViewPane\CheckPane.h"
#include "ViewPane\DropDownPane.h"
#include "ViewPane\ListPane.h"
#include "ViewPane\CollapsibleTextPane.h"
#include "ViewPane\CountedTextPane.h"
#include "ViewPane\SmartViewPane.h"

class CParentWnd;

// Buttons for CEditor
//#define CEDITOR_BUTTON_OK 0x00000001 // Duplicated from MFCMAPI.h - do not modify
#define CEDITOR_BUTTON_ACTION1 0x00000002
#define CEDITOR_BUTTON_ACTION2 0x00000004
//#define CEDITOR_BUTTON_CANCEL 0x00000008 // Duplicated from MFCMAPI.h - do not modify
#define CEDITOR_BUTTON_ACTION3 0x00000010

union ControlStruct
{
	ViewPane* lpPane;
	TextPane* lpTextPane;
	CheckPane* lpCheckPane;
	ListPane* lpListPane;
	DropDownPane* lpDropDownPane;
	CollapsibleTextPane* lpCollapsibleTextPane;
};

class CEditor : public CMyDialog
{
public:
	// Main Edit Constructor
	CEditor(
		_In_opt_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulNumFields,
		ULONG ulButtonFlags);
	CEditor(
		_In_opt_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulNumFields,
		ULONG ulButtonFlags,
		UINT uidActionButtonText1,
		UINT uidActionButtonText2,
		UINT uidActionButtonText3);
	virtual ~CEditor();

	_Check_return_ HRESULT DisplayDialog();

	// These functions can be used to set up a data editing dialog
	void SetPromptPostFix(_In_ wstring szMsg);
	void CreateControls(ULONG ulCount);
	void InitPane(ULONG iNum, ViewPane* lpPane);
	void SetStringA(ULONG i, _In_opt_z_ LPCSTR szMsg, size_t cchsz = -1);
	void SetStringW(ULONG i, _In_opt_z_ LPCWSTR szMsg, size_t cchsz = -1);
#ifdef UNICODE
#define SetString SetStringW
	//#define InitSingleLineSz InitSingleLineSzW
#else
#define SetString SetStringA
	//#define InitSingleLineSz InitSingleLineSzA
#endif
	void __cdecl SetStringf(ULONG i, _Printf_format_string_ LPCTSTR szMsg, ...);
	void LoadString(ULONG i, UINT uidMsg);
	void SetBinary(ULONG i, _In_opt_count_(cb) LPBYTE lpb, size_t cb);
	void SetHex(ULONG i, ULONG ulVal);
	void SetDecimal(ULONG i, ULONG ulVal);
	void SetSize(ULONG i, size_t cb);

	// Get values after we've done the DisplayDialog
	ViewPane* GetControl(ULONG iControl);
	LPSTR GetStringA(ULONG i);
	LPWSTR GetStringW(ULONG i);
#ifdef UNICODE
#define GetString GetStringW
#else
#define GetString GetStringA
#endif
	_Check_return_ ULONG GetHex(ULONG i);
	_Check_return_ ULONG GetDecimal(ULONG i);
	_Check_return_ ULONG GetPropTag(ULONG i);
	_Check_return_ bool GetCheck(ULONG i);
	_Check_return_ int GetDropDown(ULONG i);
	_Check_return_ DWORD_PTR GetDropDownValue(ULONG i);
	_Check_return_ HRESULT GetEntryID(ULONG i, bool bIsBase64, _Out_ size_t* lpcbBin, _Out_ LPENTRYID* lpEID);
	_Check_return_ bool GetSelectedGUID(ULONG iControl, bool bByteSwapped, _In_ LPGUID lpSelectedGUID);

	// AddIn functions
	void SetAddInTitle(_In_z_ LPWSTR szTitle);
	void SetAddInLabel(ULONG i, _In_z_ LPWSTR szLabel);

	// Use this function to implement list editing
	// return true to indicate the entry was changed, false to indicate it was not
	_Check_return_ virtual bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);

protected:
	// Functions used by derived classes during init
	void InsertColumn(ULONG ulListNum, int nCol, UINT uidText);
	void InsertColumn(ULONG ulListNum, int nCol, UINT uidText, ULONG ulPropType);
	void SetListString(ULONG iControl, ULONG iListRow, ULONG iListCol, wstring szListString);
	_Check_return_ SortListData* InsertListRow(ULONG iControl, int iRow, _In_z_ LPCTSTR szText);
	void ClearList(ULONG iControl);
	void ResizeList(ULONG uControl, bool bSort);

	// Functions used by derived classes during handle change events
	_Check_return_ CString GetStringUseControl(ULONG iControl);
	_Check_return_ ULONG GetHexUseControl(ULONG i);
	_Check_return_ ULONG GetDecimalUseControl(ULONG i);
	_Check_return_ ULONG GetPropTagUseControl(ULONG iControl);
	_Check_return_ bool GetBinaryUseControl(ULONG i, _Out_ size_t* cbBin, _Out_ LPBYTE* lpBin);
	_Check_return_ bool GetCheckUseControl(ULONG iControl);
	_Check_return_ LPSTR GetEditBoxTextA(ULONG iControl, _Out_ size_t* lpcchText = NULL);
	_Check_return_ LPWSTR GetEditBoxTextW(ULONG iControl, _Out_ size_t* lpcchText = NULL);
	_Check_return_ ULONG GetListCount(ULONG iControl);
	_Check_return_ SortListData* GetListRowData(ULONG iControl, int iRow);
	_Check_return_ bool IsDirty(ULONG iControl);

	// Called to enable/disable buttons based on number of items
	void UpdateListButtons();
	_Check_return_ bool IsValidDropDown(ULONG ulNum);
	_Check_return_ bool IsValidEdit(ULONG ulNum);
	_Check_return_ bool IsValidList(ULONG ulNum);
	_Check_return_ bool IsValidCheck(ULONG ulNum);
	BOOL OnInitDialog();
	void OnOK();
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
		ULONG ulNumFields,
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
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void SetMargins();

	// List functions and data
	_Check_return_ bool OnEditListEntry(ULONG ulListNum);
	_Check_return_ bool IsValidListWithButtons(ULONG ulNum);
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

	ControlStruct* m_lpControls; // array of controls
	ULONG m_cControls; // count of controls

	bool m_bEnableScroll;
	CWnd m_ScrollWindow;
	HWND m_hWndVertScroll;
	bool m_bScrollVisible;

	int m_iScrollClient;

	DECLARE_MESSAGE_MAP()
};