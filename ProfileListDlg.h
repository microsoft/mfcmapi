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
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPMAPITABLE lpMAPITable);
	virtual ~CProfileListDlg();

private:
	// Overrides from base class
	void OnDeleteSelectedItem();
	void OnDisplayItem();
	void OnRefreshView();

	// Menu items
	void OnAddExchangeToProfile();
	void OnAddPSTToProfile();
	void OnAddUnicodePSTToProfile();
	void OnAddServicesToMAPISVC();
	void OnAddServiceToProfile();
	void OnCreateProfile();
	void OnGetMAPISVC();
	void OnGetProfileServiceVersion();
	void OnInitMenu(CMenu* pMenu);
	void OnLaunchProfileWizard();
	void OnRemoveServicesFromMAPISVC();

	void AddPSTToProfile(BOOL bUnicodePST);

	DECLARE_MESSAGE_MAP()
};