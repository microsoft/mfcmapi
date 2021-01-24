#pragma once
#include <UI/Controls/SortList/SortListCtrl.h>
#include <core/PropertyBag/PropertyBag.h>
#include <core/model/mapiRowModel.h>
#include <core/sortlistdata/propModelData.h>

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CBaseDialog;
}

namespace controls::sortlistctrl
{
	class CSingleMAPIPropListCtrl : public CSortListCtrl
	{
	public:
		CSingleMAPIPropListCtrl(
			_In_ CWnd* pCreateParent,
			_In_ dialog::CBaseDialog* lpHostDlg,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);
		~CSingleMAPIPropListCtrl();

		// Initialization
		void SetDataSource(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ sortlistdata::sortListData* lpListData, bool bIsAB);
		void SetDataSource(const std::shared_ptr<propertybag::IMAPIPropertyBag> lpPropBag);
		void RefreshMAPIPropList();

		// Selected item accessors
		const std::shared_ptr<propertybag::IMAPIPropertyBag> GetDataSource();
		_Check_return_ std::shared_ptr<sortlistdata::propModelData> GetSelectedPropModelData() const;
		_Check_return_ bool IsModifiedPropVals() const;

		_Check_return_ bool HandleMenu(WORD wMenuSelect);
		void InitMenu(_In_ CMenu* pMenu) const;
		void SavePropsToXML();
		void OnPasteProperty();

	private:
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		void AddPropToExtraProps(ULONG ulPropTag, bool bRefresh);
		void AddPropsToExtraProps(_In_ LPSPropTagArray lpPropsToAdd, bool bRefresh);
		void FindAllNamedProps();
		void CountNamedProps();
		void LoadMAPIPropList();

		void AddPropToListBox(int iRow, const std::shared_ptr<model::mapiRowModel>& model);

		_Check_return_ bool HandleAddInMenu(WORD wMenuSelect) const;
		void OnContextMenu(_In_ CWnd* pWnd, CPoint pos);
		void OnDblclk(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void OnCopyProperty() const;
		void OnCopyTo();
		void OnDeleteProperty();
		void OnDisplayPropertyAsSecurityDescriptorPropSheet() const;
		void OnEditGivenProp(ULONG ulPropTag, const std::wstring& name);
		void OnEditGivenProperty();
		void OnEditProp();
		void OnEditPropAsRestriction(ULONG ulPropTag);
		void OnEditPropAsStream(ULONG ulType, bool bEditAsRTF);
		void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
		void OnModifyExtraProps();
		void OnOpenProperty() const;
		void OnOpenPropertyAsTable();
		void OnPasteNamedProps();

		// Custom messages
		_Check_return_ LRESULT msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);

		dialog::CBaseDialog* m_lpHostDlg{};
		bool m_bHaveEverDisplayedSomething{};
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};

		std::shared_ptr<propertybag::IMAPIPropertyBag> m_lpPropBag{};

		// Used to store prop tags added through AddPropsToExtraProps
		LPSPropTagArray m_sptExtraProps{};

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls::sortlistctrl