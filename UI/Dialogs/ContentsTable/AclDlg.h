#pragma once

#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

namespace dialog
{
	class CAclDlg : public CContentsTableDlg
	{
	public:
		CAclDlg(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ cache::CMapiObjects* lpMapiObjects,
			_In_ LPEXCHANGEMODIFYTABLE lpExchTbl,
			bool bFreeBusyVisible);
		virtual ~CAclDlg();

	private:
		// Overrides from base class
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;
		void OnDeleteSelectedItem() override;
		void OnInitMenu(_In_ CMenu* pMenu) override;
		void OnRefreshView() override;
		_Check_return_ HRESULT
		OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp) override;

		// Menu items
		void OnAddItem();
		void OnModifySelectedItem();

		_Check_return_ HRESULT GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, _In_ LPROWLIST* lppRowList) const;

		LPEXCHANGEMODIFYTABLE m_lpExchTbl;
		ULONG m_ulTableFlags;

		DECLARE_MESSAGE_MAP()
	};
}