#pragma once
// BaseDialog.h : header file

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
		CMapiObjects* lpMapiObjects,
		ULONG ulAddInContext);
	virtual ~CBaseDialog();

	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	void OnUpdateSingleMAPIPropListCtrl(LPMAPIPROP lpMAPIProp, SortListData* lpListData);
	BOOL HandleKeyDown(UINT nChar, BOOL bShift, BOOL bCtrl, BOOL bMenu);

	void UpdateTitleBarText(LPCTSTR szMsg);
	void UpdateStatusBarText(__StatusPaneEnum nPos,LPCTSTR szMsg);
	void __cdecl UpdateStatusBarText(__StatusPaneEnum nPos,UINT uidMsg,...);
	void OnOpenEntryID(LPSBinary lpBin);
	CParentWnd* GetParentWnd();
	CMapiObjects* GetMapiObjects();

protected:
	// Overrides called by child classes
	virtual void CreateDialogAndMenu(UINT nIDMenuResource);
	virtual void EnableAddInMenus(CMenu* pMenu, ULONG ulMenu, LPMENUITEM lpAddInMenu, UINT uiEnable);
	virtual BOOL HandleMenu(WORD wMenuSelect);
	virtual BOOL HandlePaste();
	virtual void OnCancel();
	virtual BOOL OnInitDialog();
	virtual void OnInitMenu(CMenu* pMenu);

	void AddMenu(UINT uiResource, UINT uidTitle, UINT uiPos);

	ULONG						m_ulAddInContext;
	ULONG						m_ulAddInMenuItems;
	BOOL						m_bIsAB;
	CSingleMAPIPropListCtrl*	m_lpPropDisplay;
	CFakeSplitter*				m_lpFakeSplitter;
	CString						m_szTitle;
	LPMAPICONTAINER				m_lpContainer;
	CMapiObjects*				m_lpMapiObjects;
	CParentWnd*					m_lpParent;

private:
	virtual BOOL HandleAddInMenu(WORD wMenuSelect);
	virtual BOOL HandleCopy();
	virtual void OnDeleteSelectedItem();
	virtual void OnEscHit();
	virtual void OnRefreshView();

	// Overrides from base class
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void	OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized);
	void	OnMenuSelect(UINT nItemID, UINT nFlags, HMENU hSysMenu);
	void	OnOK();
	void	OnSize(UINT nType, int cx, int cy);

	// Menu items
	void OnCompareEntryIDs();
	void OnComputeStoreHash();
	void OnDispatchNotifications();
	void OnEncodeID();
	void OnHelp();
	void OnHexEditor();
	void OnNotificationsOff();
	void OnNotificationsOn();
	void OnOpenMainWindow();
	void OnOptions();
	void OnOutlookVersion();

	void SetStatusWidths();

	// Custom messages
	LRESULT msgOnUpdateStatusBar(WPARAM wParam, LPARAM lParam);
	LRESULT msgOnClearSingleMAPIPropList(WPARAM wParam, LPARAM lParam);

	LONG			m_cRef;
	HICON			m_hIcon;
	CStatusBarCtrl	m_StatusBar;
	BOOL			m_bDisplayingMenuText;
	CString			m_szMenuDisplacedText;
	CAdviseSink*	m_lpBaseAdviseSink;
	ULONG_PTR		m_ulBaseAdviseConnection;
	ULONG			m_ulBaseAdviseObjectType;

	DECLARE_MESSAGE_MAP()
};