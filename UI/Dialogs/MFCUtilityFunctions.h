#pragma once
// Common functions for MFCMAPI

#include <UI/Dialogs/BaseDialog.h>

namespace dialog
{
	enum class objectType
	{
		otDefault,
		assocContents,
		contents,
		hierarchy,
		ACL,
		status,
		receive,
		rules,
		store,
		storeDeletedItems,
	};

	_Check_return_ HRESULT DisplayObject(
		_In_ LPMAPIPROP lpUnk,
		ULONG ulObjType,
		objectType tType,
		_In_opt_ CBaseDialog* lpHostDlg,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ ui::CParentWnd* lpParentWnd);

	_Check_return_ HRESULT
	DisplayObject(_In_ LPMAPIPROP lpUnk, ULONG ulObjType, objectType tType, _In_ dialog::CBaseDialog* lpHostDlg);

	_Check_return_ HRESULT DisplayExchangeTable(
		_In_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		objectType tType,
		_In_ dialog::CBaseDialog* lpHostDlg);

	_Check_return_ HRESULT
	DisplayTable(_In_ LPMAPITABLE lpTable, objectType tType, _In_ dialog::CBaseDialog* lpHostDlg);

	_Check_return_ HRESULT
	DisplayTable(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, objectType tType, _In_ dialog::CBaseDialog* lpHostDlg);

	_Check_return_ bool bShouldCancel(_In_opt_ CWnd* cWnd, HRESULT hResPrev);

	void DisplayMailboxTable(_In_ ui::CParentWnd* lpParent, _In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);
	void
	DisplayPublicFolderTable(_In_ ui::CParentWnd* lpParent, _In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);
	_Check_return_ LPMAPIFORMINFO
	ResolveMessageClass(_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects, _In_opt_ LPMAPIFOLDER lpMAPIFolder);
	_Check_return_ LPMAPIFORMINFO SelectForm(
		_In_ HWND hWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_opt_ LPMAPIFOLDER lpMAPIFolder);
} // namespace dialog