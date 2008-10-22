#pragma once
// SingleMAPIPropListCtrl.h : header file

#include "SortListCtrl.h"
class CBaseDialog;
class CMapiObjects;

class CSingleMAPIPropListCtrl : public CSortListCtrl
{
public:
	CSingleMAPIPropListCtrl(
		CWnd* pCreateParent,
		CBaseDialog* lpHostDlg,
		CMapiObjects* lpMapiObjects,
		BOOL bIsAB
		);
	virtual ~CSingleMAPIPropListCtrl();

	// Initialization
	HRESULT SetDataSource(LPMAPIPROP lpMAPIProp, SortListData* lpListData, BOOL bIsAB);
	HRESULT	RefreshMAPIPropList();

	// Selected item accessors
	ULONG			GetCountPropVals();
	LPSPropValue	GetPropVals();
	void			GetSelectedPropTag(ULONG* lpPropTag);
	BOOL			IsModifiedPropVals();

	BOOL HandleMenu(WORD wMenuSelect);
	void InitMenu(CMenu* pMenu);
	void SavePropsToXML();
	void OnPasteProperty();

private:
	HRESULT AddPropToExtraProps(ULONG ulPropTag,BOOL bRefresh);
	HRESULT AddPropsToExtraProps(LPSPropTagArray lpPropsToAdd,BOOL bRefresh);
	HRESULT FindAllNamedProps();
	HRESULT LoadMAPIPropList();
	HRESULT	SetNewProp(LPSPropValue lpNewProp);

	void AddPropToListBox(
		int iRow,
		ULONG ulPropTag,
		LPMAPINAMEID lpNameID,
		LPSPropValue lpsPropToAdd);

	BOOL HandleAddInMenu(WORD wMenuSelect);
	void OnContextMenu(CWnd *pWnd, CPoint pos);
	void OnDblclk(NMHDR* pNMHDR, LRESULT* pResult);
	void OnCopyProperty();
	void OnCopyTo();
	void OnDeleteProperty();
	void OnDisplayPropertyAsSecurityDescriptorPropSheet();
	void OnEditGivenProp(ULONG ulPropTag);
	void OnEditGivenProperty();
	void OnEditProp();
	void OnEditPropAsRestriction(ULONG ulPropTag);
	void OnEditPropAsStream(ULONG ulType, BOOL bEditAsRTF);
	void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	void OnModifyExtraProps();
	void OnOpenProperty();
	void OnOpenPropertyAsTable();
	void OnParseProperty();
	void OnPasteNamedProps();

	// Custom messages
	LRESULT	msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);

	LPMAPIPROP		m_lpMAPIProp;
	SortListData*	m_lpSourceData; // NEVER FREE THIS - It's just 'on loan' from CContentsTableListCtrl
	CBaseDialog*	m_lpHostDlg;
	CString			m_szTitle;
	BOOL			m_bHaveEverDisplayedSomething;
	BOOL			m_bIsAB;
	BOOL			m_bRowModified;
	CMapiObjects*	m_lpMapiObjects;

	// Used to store prop tags added through AddPropsToExtraProps
	LPSPropTagArray		m_sptExtraProps;

	DECLARE_MESSAGE_MAP()
};