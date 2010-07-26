#pragma once
// SingleMAPIPropListCtrl.h : header file

#include "SortListCtrl.h"
class CBaseDialog;
class CMapiObjects;

class CSingleMAPIPropListCtrl : public CSortListCtrl
{
public:
	CSingleMAPIPropListCtrl(
		_In_ CWnd* pCreateParent,
		_In_ CBaseDialog* lpHostDlg,
		_In_ CMapiObjects* lpMapiObjects,
		BOOL bIsAB
		);
	virtual ~CSingleMAPIPropListCtrl();

	// Initialization
	_Check_return_ HRESULT SetDataSource(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ SortListData* lpListData, BOOL bIsAB);
	_Check_return_ HRESULT RefreshMAPIPropList();

	// Selected item accessors
	_Check_return_ ULONG        GetCountPropVals();
	_Check_return_ LPSPropValue GetPropVals();
	void         GetSelectedPropTag(_Out_ ULONG* lpPropTag);
	_Check_return_ BOOL         IsModifiedPropVals();

	_Check_return_ BOOL HandleMenu(WORD wMenuSelect);
	void InitMenu(_In_ CMenu* pMenu);
	void SavePropsToXML();
	void OnPasteProperty();

private:
	_Check_return_ HRESULT AddPropToExtraProps(ULONG ulPropTag, BOOL bRefresh);
	_Check_return_ HRESULT AddPropsToExtraProps(_In_ LPSPropTagArray lpPropsToAdd, BOOL bRefresh);
	void FindAllNamedProps();
	void CountNamedProps();
	_Check_return_ HRESULT LoadMAPIPropList();
	_Check_return_ HRESULT SetNewProp(_In_ LPSPropValue lpNewProp);

	void AddPropToListBox(
		int iRow,
		ULONG ulPropTag,
		_In_opt_ LPMAPINAMEID lpNameID,
		_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		_In_ LPSPropValue lpsPropToAdd);

	_Check_return_ BOOL HandleAddInMenu(WORD wMenuSelect);
	void OnContextMenu(_In_ CWnd *pWnd, CPoint pos);
	void OnDblclk(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
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
	_Check_return_ LRESULT msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);

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