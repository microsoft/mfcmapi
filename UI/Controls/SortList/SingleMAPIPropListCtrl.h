#pragma once
#include <UI/Controls/SortList/SortListCtrl.h>
#include <core/PropertyBag/PropertyBag.h>

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CBaseDialog;
}

namespace controls
{
	namespace sortlistctrl
	{
		class CSingleMAPIPropListCtrl : public CSortListCtrl
		{
		public:
			CSingleMAPIPropListCtrl(
				_In_ CWnd* pCreateParent,
				_In_ dialog::CBaseDialog* lpHostDlg,
				_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
				bool bIsAB);
			virtual ~CSingleMAPIPropListCtrl();

			// Initialization
			_Check_return_ HRESULT
			SetDataSource(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ sortlistdata::sortListData* lpListData, bool bIsAB);
			_Check_return_ HRESULT SetDataSource(const std::shared_ptr<propertybag::IMAPIPropertyBag> lpPropBag, bool bIsAB);
			_Check_return_ HRESULT RefreshMAPIPropList();

			// Selected item accessors
			_Check_return_ HRESULT GetDisplayedProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray) const;
			void GetSelectedPropTag(_Out_ ULONG* lpPropTag) const;
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
			_Check_return_ HRESULT LoadMAPIPropList();

			void AddPropToListBox(
				int iRow,
				ULONG ulPropTag,
				_In_opt_ LPMAPINAMEID lpNameID,
				_In_opt_ LPSBinary
					lpMappingSignature, // optional mapping signature for object to speed named prop lookups
				_In_ LPSPropValue lpsPropToAdd);

			_Check_return_ bool HandleAddInMenu(WORD wMenuSelect) const;
			void OnContextMenu(_In_ CWnd* pWnd, CPoint pos);
			void OnDblclk(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
			void OnCopyProperty() const;
			void OnCopyTo();
			void OnDeleteProperty();
			void OnDisplayPropertyAsSecurityDescriptorPropSheet() const;
			void OnEditGivenProp(ULONG ulPropTag);
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
			bool m_bIsAB{};
			std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};

			std::shared_ptr<propertybag::IMAPIPropertyBag> m_lpPropBag{};

			// Used to store prop tags added through AddPropsToExtraProps
			LPSPropTagArray m_sptExtraProps{};

			DECLARE_MESSAGE_MAP()
		};
	} // namespace sortlistctrl
} // namespace controls