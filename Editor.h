#pragma once
// Editor.h : header file
//

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

//Buttons for CEditor
#define CEDITOR_BUTTON_OK		0x00000001
#define CEDITOR_BUTTON_ACTION1	0x00000002
#define CEDITOR_BUTTON_ACTION2	0x00000004
#define CEDITOR_BUTTON_CANCEL	0x00000008


struct DropDownStruct
{
	CComboBox	DropDown;//UI Control
	ULONG		ulDropList;//count of entries in szDropList
	UINT*		lpuidDropList;
	int			iDropValue;
};

struct CheckStruct
{
	CButton	Check;//UI Control
	BOOL	bCheckValue;
};

struct EditStruct
{
	CRichEditCtrl	EditBox;//UI Control

	size_t			cchszW;//null terminated count of lpszW since it may contain internal NULL
	LPWSTR			lpszW;
	LPSTR			lpszA;// on demand conversion of lpszW
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
	BOOL	bUseLabelControl;//whether to use an extra label control - some controls will use szLabel on their own (checkbox)
	CEdit	Label;//UI Control
	UINT	uidLabel;//Label to load
	CString	szLabel;//Text to push into UI in OnInitDialog
	UINT	nID;//id for matching change notifications back to controls
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
	//Main Edit Constructor
	CEditor(
		CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulNumFields,
		ULONG ulButtonFlags);

	~CEditor();

	HRESULT DisplayDialog();

	//These functions can be used to set up a data editing dialog
	void CreateControls(ULONG ulCount);
	void DeleteControls();
	void SetPromptPostFix(LPCTSTR szMsg);
	void InitMultiLine(ULONG i, UINT uidLabel, LPCTSTR szVal, BOOL bReadOnly);
	void InitSingleLineSz(ULONG i, UINT uidLabel, LPCTSTR szVal, BOOL bReadOnly);
	void InitSingleLine(ULONG i, UINT uidLabel, UINT uidVal, BOOL bReadOnly);
	void InitCheck(ULONG i, UINT uidLabel, BOOL bVal, BOOL bReadOnly);
	void InitList(ULONG i, UINT uidLabel, BOOL bAllowSort, BOOL bReadOnly);
	void InitDropDown(ULONG i, UINT uidLabel, ULONG ulDropList, UINT* lpuidDropList, BOOL bReadOnly);
	void InsertColumn(ULONG ulListNum,int nCol,UINT uidText);
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
	void SetSize(ULONG i, size_t cb);
	void SetHex(ULONG i, ULONG ulVal);
	void SetDecimal(ULONG i, ULONG ulVal);
	void SetDropDown(ULONG i,LPCTSTR szText);
	void ReadEditBoxIntoLPSZW(ULONG i);

	//Get values after we've done the DisplayDialog
	LPSTR	GetStringA(ULONG i);
	LPWSTR	GetStringW(ULONG i);
#ifdef UNICODE
#define GetString  GetStringW
#else
#define GetString  GetStringA
#endif
	ULONG	GetHex(ULONG i);
	ULONG	GetHexUseControl(ULONG i);
	ULONG	GetDecimal(ULONG i);
	BOOL	GetCheck(ULONG i);
	int		GetDropDown(ULONG i);
	BOOL	GetBinaryUseControl(ULONG i,size_t* cbBin,LPBYTE* lpBin);
	HRESULT GetEntryID(ULONG i, BOOL bIsBase64, size_t* cbBin, LPENTRYID* lpEID);

	// AddIn functions
	void SetAddInTitle(LPWSTR szTitle);
	void SetAddInLabel(ULONG i,LPWSTR szLabel);

protected:
	//{{AFX_MSG(CEditor)
	afx_msg void	OnSize(UINT nType, int cx, int cy);
	afx_msg void	OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	afx_msg LRESULT    OnNcHitTest(CPoint point);
	afx_msg void	OnPaint();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	afx_msg void	OnMoveListEntryDown(ULONG ulListNum);
	afx_msg void	OnAddListEntry(ULONG ulListNum);
	afx_msg BOOL	OnEditListEntry(ULONG ulListNum);
	afx_msg void	OnDeleteListEntry(ULONG ulListNum, BOOL bDoDirty);
	afx_msg void	OnMoveListEntryUp(ULONG ulListNum);
	afx_msg void	OnMoveListEntryToBottom(ULONG ulListNum);
	afx_msg void	OnMoveListEntryToTop(ULONG ulListNum);

	LPSTR GetEditBoxTextA(ULONG i);

	//use this function to implement a custom edit button
	virtual void	OnEditAction1();
	virtual void	OnEditAction2();

	//Use this function to implement list editing
	//return true to indicate the entry was changed, false to indicate it was not
	virtual BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);

	//Called to enable/disable buttons based on number of items
	virtual void	UpdateListButtons();

	BOOL	IsValidDropDown(ULONG ulNum);
	BOOL	IsValidEdit(ULONG ulNum);
	BOOL	IsValidList(ULONG ulNum);
	BOOL	IsValidListWithButtons(ULONG ulNum);
	BOOL	OnInitDialog();
	void	OnOK();

	CString					m_szTitle;
	UINT					m_uidTitle;
	UINT					m_uidPrompt;
	UINT					m_uidActionButtonText1;
	UINT					m_uidActionButtonText2;

	ControlStruct*			m_lpControls;//array of edit boxes
	ULONG					m_cControls;//count of edit boxes

	virtual ULONG			HandleChange(UINT nID);
	LRESULT					WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
private:
	HICON					m_hIcon;

	//Our UI controls. Only valid during display.
	CEdit					m_Prompt;
	CButton					m_OkButton;
	CButton					m_ActionButton1;
	CButton					m_ActionButton2;
	CButton					m_CancelButton;
	ULONG					m_cSingleLineBoxes;
	ULONG					m_cMultiLineBoxes;
	ULONG					m_cCheckBoxes;
	ULONG					m_cDropDowns;
	ULONG					m_cListBoxes;
	ULONG					m_cLabels;
	ULONG					m_cButtons;

	ULONG					m_ulListNum;//Only supporting one list right now - this is the control number for it

	//Variables that get set in the constructor
	ULONG					m_bButtonFlags;

	void	UpdateEditBoxText(ULONG i);//Only used internally - others should use SetString, etc.
	void	ClearString(ULONG i);
	void	SwapListItems(ULONG ulListNum, ULONG ulFirstItem, ULONG ulSecondItem);
	void	OnSetDefaultSize();

	int		m_iMargin;
	int		m_iButtonWidth;
	int		m_iEditHeight;
	int		m_iTextHeight;
	int		m_iButtonHeight;

	SIZE	ComputeWorkArea(SIZE sScreen);
	int		m_iMinWidth;
	int		m_iMinHeight;
	CString	m_szPromptPostFix;

	CString m_szAddInTitle;
};
