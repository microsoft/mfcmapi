#pragma once
class CSingleMAPIPropListCtrl;
class CMapiObjects;
class CFakeSplitter;
class CParentWnd;
class CAdviseSink;

#include <Dialogs/Dialog.h>
#include "enums.h"
#include "SortList/SortListData.h"

class CBaseDialog : public CMyDialog
{
public:
	CBaseDialog(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		ULONG ulAddInContext);
	virtual ~CBaseDialog();

	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	void OnUpdateSingleMAPIPropListCtrl(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ SortListData* lpListData);
	_Check_return_ bool HandleKeyDown(UINT nChar, bool bShift, bool bCtrl, bool bMenu);

	void UpdateTitleBarText(_In_ const wstring& szMsg) const;
	void UpdateTitleBarText() const;
	void UpdateStatusBarText(__StatusPaneEnum nPos, _In_ const wstring& szMsg);
	void __cdecl UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg);
	void __cdecl UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg, ULONG ulParam1);
	void __cdecl UpdateStatusBarText(__StatusPaneEnum nPos, UINT uidMsg, wstring& szParam1, wstring& szParam2, wstring& szParam3);
	void OnOpenEntryID(_In_opt_ LPSBinary lpBin);
	_Check_return_ CParentWnd* GetParentWnd() const;
	_Check_return_ CMapiObjects* GetMapiObjects() const;

	static void UpdateStatus(HWND hWndHost, __StatusPaneEnum pane, const wstring& status);

protected:
	// Overrides called by child classes
	virtual void CreateDialogAndMenu(UINT nIDMenuResource, UINT uiClassMenuResource, UINT uidClassMenuTitle);
	virtual void EnableAddInMenus(_In_ HMENU hMenu, ULONG ulMenu, _In_ LPMENUITEM lpAddInMenu, UINT uiEnable);
	_Check_return_ virtual bool HandleMenu(WORD wMenuSelect);
	_Check_return_ virtual bool HandlePaste();
	void OnCancel() override;
	BOOL OnInitDialog() override;
	void OnInitMenu(_In_opt_ CMenu* pMenu);

	ULONG m_ulAddInContext;
	ULONG m_ulAddInMenuItems;
	bool m_bIsAB;
	CSingleMAPIPropListCtrl* m_lpPropDisplay;
	CFakeSplitter* m_lpFakeSplitter;
	wstring m_szTitle;
	LPMAPICONTAINER m_lpContainer;
	CMapiObjects* m_lpMapiObjects;
	CParentWnd* m_lpParent;

private:
	_Check_return_ virtual bool HandleAddInMenu(WORD wMenuSelect);
	static void AddMenu(HMENU hMenuBar, UINT uiResource, UINT uidTitle, UINT uiPos);
	virtual void HandleCopy();
	virtual void OnDeleteSelectedItem();
	virtual void OnEscHit();
	virtual void OnRefreshView();

	// Overrides from base class
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
	void OnActivate(UINT nState, _In_ CWnd* pWndOther, BOOL bMinimized);
	void OnMenuSelect(UINT nItemID, UINT nFlags, HMENU hSysMenu);
	void OnOK() override;
	void OnSize(UINT nType, int cx, int cy);

	// Menu items
	void OnCompareEntryIDs();
	void OnComputeStoreHash();
	void OnDispatchNotifications();
	void OnHelp();
	void OnHexEditor();
	void OnNotificationsOff();
	void OnNotificationsOn();
	void OnOpenMainWindow();
	void OnOptions();
	void OnOutlookVersion();

	void SetStatusWidths();

	// Custom messages
	_Check_return_ LRESULT msgOnUpdateStatusBar(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnClearSingleMAPIPropList(WPARAM wParam, LPARAM lParam);

	LONG m_cRef;
	HICON m_hIcon;
	wstring m_StatusMessages[STATUSBARNUMPANES];
	int m_StatusWidth[STATUSBARNUMPANES];
	bool m_bDisplayingMenuText;
	wstring m_szMenuDisplacedText;
	CAdviseSink* m_lpBaseAdviseSink;
	ULONG_PTR m_ulBaseAdviseConnection;
	ULONG m_ulBaseAdviseObjectType;

	DECLARE_MESSAGE_MAP()
};