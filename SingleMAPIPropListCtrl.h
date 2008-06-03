#pragma once

#include "SortListCtrl.h"
class CBaseDialog;

/////////////////////////////////////////////////////////////////////////////
// CSingleMAPIPropListCtrl window

class CSingleMAPIPropListCtrl : public CSortListCtrl
{
	// Construction
public:
	CSingleMAPIPropListCtrl(
		CWnd* pCreateParent,
		CBaseDialog *lpHostDlg,
		BOOL bIsAB
		);
	virtual ~CSingleMAPIPropListCtrl();

	HRESULT AddPropToExtraProps(ULONG ulPropTag,BOOL bRefresh);
	HRESULT AddPropsToExtraProps(LPSPropTagArray lpPropsToAdd,BOOL bRefresh);
	HRESULT FindAllNamedProps();
	void	InitMenu(CMenu* pMenu);
	BOOL	HandleMenu(WORD wMenuSelect);
	LPSPropValue GetPropVals();
	ULONG	GetCountPropVals();
	BOOL	IsModifiedPropVals();
	void	GetSelectedPropTag(ULONG* lpPropTag);
	void	SavePropsToXML();
	HRESULT SetDataSource(LPMAPIPROP lpMAPIProp, SortListData* lpListData, BOOL bIsAB);

	void	OnDeleteProperty();
	void	OnDisplayPropertyAsSecurityDescriptorPropSheet();
	void	OnParseProperty();
	void	OnEditProp();
	void	OnEditGivenProp(ULONG ulPropTag);
	void	OnEditPropAsStream(ULONG ulType, BOOL bEditAsRTF);
	void	OnEditPropAsRestriction(ULONG ulPropTag);
	void	OnOpenProperty();
	void	OnCopyProperty();
	void	OnPasteProperty();
	void	OnCopyTo();
	void	OnModifyExtraProps();
	void	OnEditGivenProperty();
	void	OnOpenPropertyAsTable();
	void	OnPasteNamedProps();
	HRESULT	RefreshMAPIPropList();

	// Generated message map functions
protected:
	//{{AFX_MSG(CSingleMAPIPropListCtrl)
	afx_msg void OnContextMenu(CWnd *pWnd, CPoint pos);
	afx_msg void OnDblclk(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg LRESULT	msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
private:
	LPMAPIPROP		m_lpMAPIProp;
	SortListData*	m_lpSourceData;//NEVER FREE THIS - It's just 'on loan' from CContentsTableListCtrl
	CBaseDialog*	m_lpHostDlg;
	CString			m_szTitle;
	BOOL			m_bHaveEverDisplayedSomething;
	BOOL			m_bIsAB;
	BOOL			m_bRowModified;

	void AddPropToListBox(
		int iRow,
		ULONG ulPropTag,
		LPMAPINAMEID lpNameID,
		LPSPropValue lpsPropToAdd);
	HRESULT LoadMAPIPropList();
	HRESULT	SetNewProp(LPSPropValue lpNewProp);

//Used to store prop tags added through AddPropsToExtraProps
	LPSPropTagArray		m_sptExtraProps;

	BOOL HandleAddInMenu(WORD wMenuSelect);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
