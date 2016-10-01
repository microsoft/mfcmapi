#pragma once
// ProfileListDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CProfileListDlg : public CContentsTableDlg
{
public:
	CProfileListDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPMAPITABLE lpMAPITable);
	virtual ~CProfileListDlg();

private:
	// Overrides from base class
	void OnDeleteSelectedItem();
	void OnDisplayItem();
	void OnRefreshView();
	void HandleCopy();
	_Check_return_ bool HandlePaste();

	// Menu items
	void OnAddExchangeToProfile();
	void OnAddPSTToProfile();
	void OnAddUnicodePSTToProfile();
	void OnAddServicesToMAPISVC();
	void OnAddServiceToProfile();
	void OnCreateProfile();
	void OnGetMAPISVC();
	void OnGetProfileServiceVersion();
	void OnInitMenu(_In_ CMenu* pMenu);
	void OnLaunchProfileWizard();
	void OnRemoveServicesFromMAPISVC();
	void OnSetDefaultProfile();
	void OnOpenProfileByName();
	void OnExportProfile();

	void AddPSTToProfile(bool bUnicodePST);

	DECLARE_MESSAGE_MAP()
};