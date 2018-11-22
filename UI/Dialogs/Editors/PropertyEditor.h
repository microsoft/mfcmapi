#pragma once
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/ViewPane/SmartViewPane.h>

namespace dialog
{
	namespace editor
	{
		_Check_return_ HRESULT DisplayPropertyEditor(
			_In_ CWnd* pParentWnd,
			UINT uidTitle,
			UINT uidPrompt,
			bool bIsAB,
			_In_opt_ LPVOID lpAllocParent,
			_In_opt_ LPMAPIPROP lpMAPIProp,
			ULONG ulPropTag,
			bool bMVRow,
			_In_opt_ const _SPropValue* lpsPropValue,
			_Inout_opt_ LPSPropValue* lpNewValue);

		class CPropertyEditor : public CEditor
		{
		public:
			CPropertyEditor(
				_In_ CWnd* pParentWnd,
				UINT uidTitle,
				UINT uidPrompt,
				bool bIsAB,
				bool bMVRow,
				_In_opt_ LPVOID lpAllocParent,
				_In_opt_ LPMAPIPROP lpMAPIProp,
				ULONG ulPropTag,
				_In_opt_ const _SPropValue* lpsPropValue);
			virtual ~CPropertyEditor();

			// Get values after we've done the DisplayDialog
			_Check_return_ LPSPropValue DetachModifiedSPropValue();

		private:
			BOOL OnInitDialog() override;
			void InitPropertyControls();
			void WriteStringsToSPropValue();
			void WriteSPropValueToObject() const;
			_Check_return_ ULONG HandleChange(UINT nID) override;
			void OnOK() override;
			void UpdateParser(const std::vector<BYTE>& bin) const;

			// source variables
			LPMAPIPROP m_lpMAPIProp;
			ULONG m_ulPropTag;
			bool m_bIsAB; // whether the tag is from the AB or not
			const _SPropValue* m_lpsInputValue;
			LPSPropValue m_lpsOutputValue;
			bool m_bDirty;
			bool m_bMVRow; // whether this row came from a multivalued property. Used for smart view parsing.
			viewpane::SmartViewPane* m_lpSmartView;

			// all calls to MAPIAllocateMore will use m_lpAllocParent
			// this is not something to be freed
			LPVOID m_lpAllocParent;
		};
	} // namespace editor
} // namespace dialog