#pragma once
#include "SortListCtrl.h"
#include "PropertyBag/PropertyBag.h"

class CBaseDialog;
class CMapiObjects;

class CSingleMAPIPropListCtrl : public CSortListCtrl
{
public:
	CSingleMAPIPropListCtrl(
		_In_ CWnd* pCreateParent,
		_In_ CBaseDialog* lpHostDlg,
		_In_ CMapiObjects* lpMapiObjects,
		bool bIsAB
	);
	virtual ~CSingleMAPIPropListCtrl();

	// Initialization
	_Check_return_ HRESULT SetDataSource(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ SortListData* lpListData, bool bIsAB);
	_Check_return_ HRESULT SetDataSource(_In_opt_ LPMAPIPROPERTYBAG lpPropBag, bool bIsAB);
	_Check_return_ HRESULT RefreshMAPIPropList();

	// Selected item accessors
	_Check_return_ HRESULT GetDisplayedProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray);
	void GetSelectedPropTag(_Out_ ULONG* lpPropTag);
	_Check_return_ bool IsModifiedPropVals();

	_Check_return_ bool HandleMenu(WORD wMenuSelect);
	void InitMenu(_In_ CMenu* pMenu);
	void SavePropsToXML();
	void OnPasteProperty();

private:
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	_Check_return_ HRESULT AddPropToExtraProps(ULONG ulPropTag, bool bRefresh);
	_Check_return_ HRESULT AddPropsToExtraProps(_In_ LPSPropTagArray lpPropsToAdd, bool bRefresh);
	void FindAllNamedProps();
	void CountNamedProps();
	_Check_return_ HRESULT LoadMAPIPropList();

	void AddPropToListBox(
		int iRow,
		ULONG ulPropTag,
		_In_opt_ LPMAPINAMEID lpNameID,
		_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		_In_ LPSPropValue lpsPropToAdd);

	_Check_return_ bool HandleAddInMenu(WORD wMenuSelect);
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
	void OnEditPropAsStream(ULONG ulType, bool bEditAsRTF);
	void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	void OnModifyExtraProps();
	void OnOpenProperty();
	void OnOpenPropertyAsTable();
	void OnPasteNamedProps();

	// Custom messages
	_Check_return_ LRESULT msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);

	CBaseDialog* m_lpHostDlg;
	bool m_bHaveEverDisplayedSomething;
	bool m_bIsAB;
	CMapiObjects* m_lpMapiObjects;

	LPMAPIPROPERTYBAG m_lpPropBag;

	// Used to store prop tags added through AddPropsToExtraProps
	LPSPropTagArray m_sptExtraProps;

	DECLARE_MESSAGE_MAP()
};