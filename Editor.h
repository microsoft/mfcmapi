#pragma once
// Editor.h : Generic edit dialog built on CDialog

class CParentWnd;

#include "enums.h"
#include "SortListCtrl.h"

enum __ControlTypes
{
	CTRL_EDIT,
	CTRL_CHECK,
	CTRL_LIST,
	CTRL_DROPDOWN
};

// Buttons for CEditor
#define CEDITOR_BUTTON_OK		0x00000001
#define CEDITOR_BUTTON_ACTION1	0x00000002
#define CEDITOR_BUTTON_ACTION2	0x00000004
#define CEDITOR_BUTTON_CANCEL	0x00000008

struct DropDownStruct
{
	CComboBox	DropDown; // UI Control
	ULONG		ulDropList; // count of entries in szDropList
	UINT*		lpuidDropList;
	int			iDropSelection;
	DWORD_PTR	iDropSelectionValue;
};

struct CheckStruct
{
	CButton	Check; // UI Control
	BOOL	bCheckValue;
};

struct EditStruct
{
	CRichEditCtrl	EditBox; // UI Control

	LPWSTR			lpszW;
	LPSTR			lpszA; // on demand conversion of lpszW
	size_t			cchsz; // length of string - maintained to preserve possible internal NULLs, includes NULL terminator
	BOOL			bMultiline;
};

struct __ListButtons
{
	UINT	uiButtonID;
};
#define NUMLISTBUTTONS 7

struct ListStruct
{
	CSortListCtrl	List;

	CButton			ButtonArray[NUMLISTBUTTONS];
	BOOL			bDirty;
	BOOL			bAllowSort;
};

struct ControlStruct
{
	ULONG	ulCtrlType;
	BOOL	bReadOnly;
	BOOL	bUseLabelControl; // whether to use an extra label control - some controls will use szLabel on their own (checkbox)
	CEdit	Label; // UI Control
	UINT	uidLabel; // Label to load
	CString	szLabel; // Text to push into UI in OnInitDialog
	UINT	nID; // id for matching change notifications back to controls
	union
	{
		CheckStruct*	lpCheck;
		EditStruct*		lpEdit;
		ListStruct*		lpList;
		DropDownStruct* lpDropDown;
	} UI;
};

class CEditor : public CDialog
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
		_In_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulNumFields,
		ULONG ulButtonFlags,
		UINT uidActionButtonText1,
		UINT uidActionButtonText2);
	virtual ~CEditor();

	_Check_return_ HRESULT DisplayDialog();

	// These functions can be used to set up a data editing dialog
	void CreateControls(ULONG ulCount);
	void SetPromptPostFix(_In_opt_z_ LPCTSTR szMsg);
	void InitMultiLine(ULONG i, UINT uidLabel, _In_opt_z_ LPCTSTR szVal, BOOL bReadOnly);
	void InitSingleLineSz(ULONG i, UINT uidLabel, _In_opt_z_ LPCTSTR szVal, BOOL bReadOnly);
	void InitSingleLine(ULONG i, UINT uidLabel, UINT uidVal, BOOL bReadOnly);
	void InitCheck(ULONG i, UINT uidLabel, BOOL bVal, BOOL bReadOnly);
	void InitDropDown(ULONG i, UINT uidLabel, ULONG ulDropList, _In_opt_count_(ulDropList) UINT* lpuidDropList, BOOL bReadOnly);
	void SetStringA(ULONG i, _In_opt_z_ LPCSTR szMsg, size_t cchsz = -1);
	void SetStringW(ULONG i, _In_opt_z_ LPCWSTR szMsg, size_t cchsz = -1);
#ifdef UNICODE
#define SetString  SetStringW
#else
#define SetString  SetStringA
#endif
	void __cdecl SetStringf(ULONG i, _Printf_format_string_ LPCTSTR szMsg, ...);
	void LoadString(ULONG i, UINT uidMsg);
	void SetBinary(ULONG i, _In_opt_count_(cb) LPBYTE lpb, size_t cb);
	void SetHex(ULONG i, ULONG ulVal);
	void SetDecimal(ULONG i, ULONG ulVal);

	// Get values after we've done the DisplayDialog
	LPSTR  GetStringA(ULONG i);
	LPWSTR GetStringW(ULONG i);
#ifdef UNICODE
#define GetString  GetStringW
#else
#define GetString  GetStringA
#endif
	_Check_return_ ULONG GetHex(ULONG i);
	_Check_return_ ULONG GetDecimal(ULONG i);
	_Check_return_ BOOL  GetCheck(ULONG i);
	_Check_return_ int   GetDropDown(ULONG i);
	_Check_return_ DWORD_PTR GetDropDownValue(ULONG i);
	_Check_return_ HRESULT GetEntryID(ULONG i, BOOL bIsBase64, _Out_ size_t* cbBin, _Out_ LPENTRYID* lpEID);

	// AddIn functions
	void SetAddInTitle(_In_z_ LPWSTR szTitle);
	void SetAddInLabel(ULONG i, _In_z_ LPWSTR szLabel);

protected:
	// Functions used by derived classes during init
	void InitList(ULONG i, UINT uidLabel, BOOL bAllowSort, BOOL bReadOnly);
	void InitEditFromBinaryStream(ULONG iControl, _In_ LPSTREAM lpStreamIn);
	void InsertColumn(ULONG ulListNum, int nCol, UINT uidText);
	void SetSize(ULONG i, size_t cb);
	void SetDropDownSelection(ULONG i, _In_opt_z_ LPCTSTR szText);
	void SetListString(ULONG iControl, ULONG iListRow, ULONG iListCol, _In_opt_z_ LPCTSTR szListString);
	_Check_return_ SortListData* InsertListRow(ULONG iControl, int iRow, _In_z_ LPCTSTR szText);
	void InsertDropString(ULONG iControl, int iRow, _In_z_ LPCTSTR szText);
	void SetEditReadOnly(ULONG iControl);
	void ClearList(ULONG iControl);
	void ResizeList(ULONG uControl, BOOL bSort);

	// Functions used by derived classes during handle change events
	_Check_return_ CString GetDropStringUseControl(ULONG iControl);
	_Check_return_ CString GetStringUseControl(ULONG iControl);
	_Check_return_ ULONG   GetHexUseControl(ULONG i);
	_Check_return_ BOOL    GetBinaryUseControl(ULONG i, _Out_ size_t* cbBin, _Out_ LPBYTE* lpBin);
	_Check_return_ BOOL    GetCheckUseControl(ULONG iControl);
	_Check_return_ LPSTR   GetEditBoxTextA(ULONG iControl, _Out_ size_t* lpcchText = NULL);
	_Check_return_ LPWSTR  GetEditBoxTextW(ULONG iControl, _Out_ size_t* lpcchText = NULL);
	_Check_return_ int     GetDropDownSelection(ULONG iControl);
	_Check_return_ DWORD_PTR GetDropDownSelectionValue(ULONG iControl);
	_Check_return_ ULONG   GetListCount(ULONG iControl);
	_Check_return_ SortListData* GetListRowData(ULONG iControl, int iRow);
	_Check_return_ SortListData* GetSelectedListRowData(ULONG iControl);
	_Check_return_ BOOL    ListDirty(ULONG iControl);
	_Check_return_ BOOL    EditDirty(ULONG iControl);
	void    AppendString(ULONG i, _In_z_ LPCTSTR szMsg);
	void    ClearView(ULONG i);

	// Called to enable/disable buttons based on number of items
	void UpdateListButtons();
	_Check_return_ BOOL IsValidDropDown(ULONG ulNum);
	_Check_return_ BOOL IsValidEdit(ULONG ulNum);
	_Check_return_ BOOL IsValidList(ULONG ulNum);
	_Check_return_ BOOL IsValidCheck(ULONG iControl);
	_Check_return_ BOOL OnInitDialog();
	void OnOK();
	void OnContextMenu(_In_ CWnd *pWnd, CPoint pos);

	// protected since derived classes need to call the base implementation
	_Check_return_ virtual ULONG HandleChange(UINT nID);

private:
	// Overridable functions
	// use these functions to implement custom edit buttons
	virtual void OnEditAction1();
	virtual void OnEditAction2();

	// Use this function to implement list editing
	// return true to indicate the entry was changed, false to indicate it was not
	_Check_return_ virtual BOOL DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);

	// private constructor invoked from public constructors
	void Constructor(
		_In_opt_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulNumFields,
		ULONG ulButtonFlags,
		UINT uidActionButtonText1,
		UINT uidActionButtonText2);

	void    DeleteControls();
	void    OnSize(UINT nType, int cx, int cy);
	void    OnGetMinMaxInfo(_Inout_ MINMAXINFO* lpMMI);
	void    OnSetDefaultSize();
	_Check_return_ SIZE    ComputeWorkArea(SIZE sScreen);
	_Check_return_ LRESULT OnNcHitTest(CPoint point);
	void    OnPaint();
	_Check_return_ LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

	void    GetEditBoxText(ULONG i);
	void    SetEditBoxText(ULONG i); // Only used internally - others should use SetString, etc.
	void    ClearString(ULONG i);
	void    SwapListItems(ULONG ulListNum, ULONG ulFirstItem, ULONG ulSecondItem);

	// List functions and data
	void    OnMoveListEntryDown(ULONG ulListNum);
	void    OnAddListEntry(ULONG ulListNum);
	_Check_return_ BOOL    OnEditListEntry(ULONG ulListNum);
	void    OnDeleteListEntry(ULONG ulListNum, BOOL bDoDirty);
	void    OnMoveListEntryUp(ULONG ulListNum);
	void    OnMoveListEntryToBottom(ULONG ulListNum);
	void    OnMoveListEntryToTop(ULONG ulListNum);
	_Check_return_ BOOL    IsValidListWithButtons(ULONG ulNum);
	ULONG	m_ulListNum; // Only supporting one list right now - this is the control number for it

	// Our UI controls. Only valid during display.
	CEdit	m_Prompt;
	CButton	m_OkButton;
	CButton	m_ActionButton1;
	CButton	m_ActionButton2;
	CButton	m_CancelButton;
	ULONG	m_cSingleLineBoxes;
	ULONG	m_cMultiLineBoxes;
	ULONG	m_cCheckBoxes;
	ULONG	m_cDropDowns;
	ULONG	m_cListBoxes;
	ULONG	m_cLabels;
	ULONG	m_cButtons;

	// Variables that get set in the constructor
	ULONG	m_bButtonFlags;

	// Size calculations
	int		m_iMargin;
	int		m_iButtonWidth;
	int		m_iEditHeight;
	int		m_iTextHeight;
	int		m_iButtonHeight;
	int		m_iMinWidth;
	int		m_iMinHeight;

	// Title bar, prompt and icon
	UINT	m_uidTitle;
	CString	m_szTitle;
	UINT	m_uidPrompt;
	CString	m_szPromptPostFix;
	CString m_szAddInTitle;
	HICON	m_hIcon;

	UINT	m_uidActionButtonText1;
	UINT	m_uidActionButtonText2;

	ControlStruct*	m_lpControls; // array of controls
	ULONG			m_cControls; // count of controls

	DECLARE_MESSAGE_MAP()
};

void CleanString(_In_ CString* lpString);
void CleanHexString(_In_ CString* lpHexString);