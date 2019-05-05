#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog
{
	namespace editor
	{
		class CPropertySelector : public CEditor
		{
		public:
			CPropertySelector(bool bIncludeABProps, _In_ LPMAPIPROP lpMAPIProp, _In_ CWnd* pParentWnd);
			virtual ~CPropertySelector();

			_Check_return_ ULONG GetPropertyTag() const;
			_Check_return_ bool
			DoListEdit(ULONG ulListNum, int iItem, _In_ controls::sortlistdata::SortListData* lpData) override;

		private:
			BOOL OnInitDialog() override;
			void OnOK() override;
			_Check_return_ controls::sortlistdata::SortListData* GetSelectedListRowData(ULONG id) const;

			ULONG m_ulPropTag;
			bool m_bIncludeABProps;
			LPMAPIPROP m_lpMAPIProp;
		};
	} // namespace editor
} // namespace dialog