#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

namespace dialog
{
	class CProviderTableDlg : public CContentsTableDlg
	{
	public:
		CProviderTableDlg(
			_In_ CParentWnd* pParentWnd,
			_In_ CMapiObjects* lpMapiObjects,
			_In_ LPMAPITABLE lpMAPITable,
			_In_ LPPROVIDERADMIN lpProviderAdmin);
		virtual ~CProviderTableDlg();

	private:
		// Overrides from base class
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;
		_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp) override;

		// Menu items
		void OnOpenProfileSection();

		LPPROVIDERADMIN m_lpProviderAdmin;

		DECLARE_MESSAGE_MAP()
	};
}