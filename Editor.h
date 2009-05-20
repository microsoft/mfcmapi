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
		CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulNumFields,
		ULONG ulButtonFlags);
	CEditor(
		CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulNumFields,
		ULONG ulButtonFlags,
		UINT uidActionButtonText1,
		UINT uidActionButtonText2);
	virtual ~CEditor();

	HRESULT DisplayDialog();

	// These functions can be used to set up a data editing dialog
	void CreateControls(ULONG ulCount);
	void SetPromptPostFix(LPCTSTR szMsg);
	void InitMultiLine(ULONG i, UINT uidLabel, LPCTSTR szVal, BOOL bReadOnly);
	void InitSingleLineSz(ULONG i, UINT uidLabel, LPCTSTR szVal, BOOL bReadOnly);
	void InitSingleLine(ULONG i, UINT uidLabel, UINT uidVal, BOOL bReadOnly);
	void InitCheck(ULONG i, UINT uidLabel, BOOL bVal, BOOL bReadOnly);
	void InitDropDown(ULONG i, UINT uidLabel, ULONG ulDropList, UINT* lpuidDropList, BOOL bReadOnly);
	void SetStringA(ULONG i,LPCSTR szMsg);
	void SetStringW(ULONG i,LPCWSTR szMsg);
#ifdef UNICODE
#define SetString  SetStringW
#else
#define SetString  SetStringA
#endif
	void __cdecl SetStringf(ULONG i,LPCTSTR szMsg,...);
	void LoadString(ULONG i, UINT uidMsg);
	void SetBinary(ULONG i,LPBYTE lpb, size_t cb);
	void SetHex(ULONG i, ULONG ulVal);
	void SetDecimal(ULONG i, ULONG ulVal);

	// Get values after we've done the DisplayDialog
	LPSTR	GetStringA(ULONG i);
	LPWSTR	GetStringW(ULONG i);
#ifdef UNICODE
#define GetString  GetStringW
#else
#define GetString  GetStringA
#endif
	ULONG	GetHex(ULONG i);
	ULONG	GetDecimal(ULONG i);
	BOOL	GetCheck(ULONG i);
	int		GetDropDown(ULONG i);
	DWORD_PTR GetDropDownValue(ULONG i);
	HRESULT GetEntryID(ULONG i, BOOL bIsBase64, size_t* cbBin, LPENTRYID* lpEID);

	// AddIn functions
	void SetAddInTitle(LPWSTR szTitle);
	void SetAddInLabel(ULONG i,LPWSTR szLabel);

protected:
	// Functions used by derived classes during init
	void InitList(ULONG i, UINT uidLabel, BOOL bAllowSort, BOOL bReadOnly);
	void InitEditFromStream(ULONG iControl, LPSTREAM lpStreamIn, BOOL bUnicode, BOOL bRTF);
	void InsertColumn(ULONG ulListNum,int nCol,UINT uidText);
	void SetSize(ULONG i, size_t cb);
	void SetDropDownSelection(ULONG i,LPCTSTR szText);
	void SetListString(ULONG iControl, ULONG iListRow, ULONG iListCol, LPCTSTR szListString);
	SortListData* InsertListRow(ULONG iControl, int iRow, LPCTSTR szText);
	void InsertDropString(ULONG iControl, int iRow, LPCTSTR szText);
	void SetEditReadOnly(ULONG iControl);
	void ClearList(ULONG iControl);
	void ResizeList(ULONG uControl, BOOL bSort);

	// Functions used by derived classes during handle change events
	CString	GetDropStringUseControl(ULONG iControl);
	CString	GetStringUseControl(ULONG iControl);
	ULONG	GetHexUseControl(ULONG i);
	BOOL	GetBinaryUseControl(ULONG i,size_t* cbBin,LPBYTE* lpBin);
	BOOL	GetCheckUseControl(ULONG iControl);
	LPSTR	GetEditBoxTextA(ULONG i);
	LPWSTR	GetEditBoxTextW(ULONG iControl);
	void	GetEditBoxStream(ULONG iControl, LPSTREAM lpStreamOut, BOOL bUnicode, BOOL bRTF);
	int		GetDropDownSelection(ULONG iControl);
	DWORD_PTR GetDropDownSelectionValue(ULONG iControl);
	ULONG	GetListCount(ULONG iControl);
	SortListData* GetListRowData(ULONG iControl, int iRow);
	SortListData* GetSelectedListRowData(ULONG iControl);
	BOOL	ListDirty(ULONG iControl);
	BOOL	EditDirty(ULONG iControl);

	// Called to enable/disable buttons based on number of items
	void	UpdateListButtons();
	BOOL	IsValidDropDown(ULONG ulNum);
	BOOL	IsValidEdit(ULONG ulNum);
	BOOL	IsValidList(ULONG ulNum);
	BOOL	IsValidCheck(ULONG iControl);
	BOOL	OnInitDialog();
	void	OnOK();

	// protected since derived classes need to call the base implementation
	virtual ULONG	HandleChange(UINT nID);

private:
	// Overridable functions
	// use these functions to implement custom edit buttons
	virtual void	OnEditAction1();
	virtual void	OnEditAction2();

	// Use this function to implement list editing
	// return true to indicate the entry was changed, false to indicate it was not
	virtual BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);

	// private constructor invoked from public constructors
	void Constructor(
		CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulNumFields,
		ULONG ulButtonFlags,
		UINT uidActionButtonText1,
		UINT uidActionButtonText2);

	void	DeleteControls();
	void	OnSize(UINT nType, int cx, int cy);
	void	OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	void	OnSetDefaultSize();
	SIZE	ComputeWorkArea(SIZE sScreen);
	LRESULT	OnNcHitTest(CPoint point);
	void	OnPaint();
	LRESULT	WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

	void	GetEditBoxText(ULONG i);
	void	SetEditBoxText(ULONG i); // Only used internally - others should use SetString, etc.
	void	ClearString(ULONG i);
	void	SwapListItems(ULONG ulListNum, ULONG ulFirstItem, ULONG ulSecondItem);

	// List functions and data
	void	OnMoveListEntryDown(ULONG ulListNum);
	void	OnAddListEntry(ULONG ulListNum);
	BOOL	OnEditListEntry(ULONG ulListNum);
	void	OnDeleteListEntry(ULONG ulListNum, BOOL bDoDirty);
	void	OnMoveListEntryUp(ULONG ulListNum);
	void	OnMoveListEntryToBottom(ULONG ulListNum);
	void	OnMoveListEntryToTop(ULONG ulListNum);
	BOOL	IsValidListWithButtons(ULONG ulNum);
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