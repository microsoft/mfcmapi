#pragma once
// BaseDialog.h : header file
//

class CSingleMAPIPropListCtrl;
class CMapiObjects;
class CFakeSplitter;
class CParentWnd;
class CAdviseSink;

#include "enums.h"

class CBaseDialog : public CDialog
{
public:
	CBaseDialog(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		ULONG ulAddInContext);
	~CBaseDialog();

	virtual STDMETHODIMP_(ULONG) AddRef()
	{
		LONG lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(_T("CBaseDialog"),lCount);// STRING_OK
		DebugPrint(DBGRefCount,_T("CBaseDialog::AddRef(\"%s\")\n"),m_szTitle);
		return m_cRef;
	}

	virtual STDMETHODIMP_(ULONG) Release()
	{
		LONG lCount = InterlockedDecrement(&m_cRef);
		DebugPrint(DBGRefCount,_T("CBaseDialog::Release(\"%s\")\n"),m_szTitle);
		TRACE_RELEASE(_T("CBaseDialog"),lCount);// STRING_OK
		if (!lCount) delete this;
		return lCount;
	}

	virtual void OnUpdateSingleMAPIPropListCtrl(LPMAPIPROP lpMAPIProp, SortListData* lpListData);
	void UpdateStatusBarText(__StatusPaneEnum nPos,LPCTSTR szMsg);
	void __cdecl UpdateStatusBarText(__StatusPaneEnum nPos,UINT uidMsg,...);
	void UpdateTitleBarText(LPCTSTR szMsg);
	virtual BOOL HandleKeyDown(UINT nChar, BOOL bShift, BOOL bCtrl, BOOL bMenu);
	afx_msg void OnCompareEntryIDs();
	afx_msg void OnOpenEntryID(LPSBinary lpBin);
	afx_msg void OnComputeStoreHash();
	afx_msg void OnEncodeID();

	void OnHexEditor();

	CMapiObjects*			m_lpMapiObjects;
	CParentWnd*				m_lpParent;

protected:
	virtual BOOL			OnInitDialog();
	afx_msg virtual void	OnInitMenu(CMenu* pMenu);

private:
	//{{AFX_MSG(CBaseDialog)
	afx_msg void			OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized);
	afx_msg void			OnSize(UINT nType, int cx, int cy);
	afx_msg void			OnMenuSelect(UINT nItemID, UINT nFlags, HMENU hSysMenu);
	afx_msg void			OnNotificationsOn();
	afx_msg void			OnNotificationsOff();
	afx_msg void			OnDispatchNotifications();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	BOOL CreateDialogAndMenu(UINT nIDMenuResource);
	BOOL AddMenu(UINT uiResource, UINT uidTitle, UINT uiPos);

	virtual void	OnRefreshView();
	virtual void	OnDeleteSelectedItem();
	virtual BOOL	HandleCopy();
	virtual BOOL	HandlePaste();
	virtual void	OnEscHit();
	void	OnCancel();
	void	OnOptions();
	void	OnOpenMainWindow();
	void	OnHelp();
	void	OnOK();

	ULONG m_ulAddInContext;
	ULONG m_ulAddInMenuItems;
	virtual void EnableAddInMenus(CMenu* pMenu, ULONG ulMenu, LPMENUITEM lpAddInMenu, UINT uiEnable);
	virtual BOOL HandleAddInMenu(WORD wMenuSelect);

	void SetStatusWidths();

	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	virtual BOOL	HandleMenu(WORD wMenuSelect);

	//Custom messages
	afx_msg LRESULT		msgOnUpdateStatusBar(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnClearSingleMAPIPropList(WPARAM wParam, LPARAM lParam);

	LPMAPICONTAINER			m_lpContainer;
	BOOL					m_bIsAB;
	CSingleMAPIPropListCtrl*	m_lpPropDisplay;
	LONG					m_cRef;
	HICON					m_hIcon;
	CStatusBarCtrl			m_StatusBar;
	CFakeSplitter*			m_lpFakeSplitter;
	CString					m_szTitle;
	BOOL					m_bDisplayingMenuText;
	CString					m_szMenuDisplacedText;

	CAdviseSink*			m_lpBaseAdviseSink;
	ULONG_PTR				m_ulBaseAdviseConnection;
	ULONG					m_ulBaseAdviseObjectType;
};
