#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	class CPropertySelector : public CEditor
	{
	public:
		CPropertySelector(bool bIncludeABProps, _In_ LPMAPIPROP lpMAPIProp, _In_ CWnd* pParentWnd);
		~CPropertySelector();

		_Check_return_ ULONG GetPropertyTag() const;
		_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData) override;

	private:
		BOOL OnInitDialog() override;
		void OnOK() override;
		_Check_return_ sortlistdata::sortListData* GetSelectedListRowData(ULONG id) const;

		ULONG m_ulPropTag;
		bool m_bIncludeABProps;
		LPMAPIPROP m_lpMAPIProp;
	};
} // namespace dialog::editor