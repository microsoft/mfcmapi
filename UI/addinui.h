#pragma once
#include <core/addin/mfcmapi.h>

namespace ui::addinui
{
	EXTERN_C _Check_return_ __declspec(dllexport) HRESULT
		__cdecl SimpleDialog(_In_z_ LPWSTR szTitle, _Printf_format_string_ LPWSTR szMsg, ...);
	EXTERN_C _Check_return_ __declspec(dllexport) HRESULT
		__cdecl ComplexDialog(_In_ LPADDINDIALOG lpDialog, _Out_opt_ LPADDINDIALOGRESULT* lppDialogResult);
	EXTERN_C __declspec(dllexport) void __cdecl FreeDialogResult(_In_ LPADDINDIALOGRESULT lpDialogResult);

	_Check_return_ ULONG ExtendAddInMenu(HMENU hMenu, ULONG ulAddInContext);
	_Check_return_ LPMENUITEM GetAddinMenuItem(HWND hWnd, UINT uidMsg) noexcept;
	void InvokeAddInMenu(_In_opt_ LPADDINMENUPARAMS lpParams);
} // namespace ui::addinui