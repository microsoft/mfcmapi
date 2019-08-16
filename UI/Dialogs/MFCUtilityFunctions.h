#pragma once
// Common functions for MFC MAPI

#include <UI/Dialogs/BaseDialog.h>

namespace dialog
{
	enum ObjectType
	{
		otDefault,
		otAssocContents,
		otContents,
		otHierarchy,
		otACL,
		otStatus,
		otReceive,
		otRules,
		otStore,
		otStoreDeletedItems,
	};

	_Check_return_ HRESULT
	DisplayObject(_In_ LPMAPIPROP lpUnk, ULONG ulObjType, ObjectType tType, _In_ dialog::CBaseDialog* lpHostDlg);

	_Check_return_ HRESULT DisplayExchangeTable(
		_In_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		ObjectType tType,
		_In_ dialog::CBaseDialog* lpHostDlg);

	_Check_return_ HRESULT
	DisplayTable(_In_ LPMAPITABLE lpTable, ObjectType tType, _In_ dialog::CBaseDialog* lpHostDlg);

	_Check_return_ HRESULT
	DisplayTable(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, ObjectType tType, _In_ dialog::CBaseDialog* lpHostDlg);

	_Check_return_ bool bShouldCancel(_In_opt_ CWnd* cWnd, HRESULT hResPrev);

	void DisplayMailboxTable(_In_ ui::CParentWnd* lpParent, _In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);
	void DisplayPublicFolderTable(_In_ ui::CParentWnd* lpParent, _In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);
	void ResolveMessageClass(
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_opt_ LPMAPIFOLDER lpMAPIFolder,
		_Out_ LPMAPIFORMINFO* lppMAPIFormInfo);
	void SelectForm(
		_In_ HWND hWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_opt_ LPMAPIFOLDER lpMAPIFolder,
		_Out_ LPMAPIFORMINFO* lppMAPIFormInfo);
} // namespace dialog