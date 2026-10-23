#pragma once
#include <core/mapi/extraPropTags.h>

namespace ui::mapiui
{
	void initCallbacks();

	LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB);

	void ExportMessages(_In_ LPMAPIFOLDER lpFolder, HWND hWnd);
	_Check_return_ HRESULT WriteAttachmentToFile(_In_ LPATTACH lpAttach, HWND hWnd);

	_Check_return_ HRESULT GetConversionToEMLOptions(
		_In_ CWnd* pParentWnd,
		_Out_opt_ CCSFLAGS* lpConvertFlags,
		_Out_opt_ ENCODINGTYPE* lpet,
		_Out_opt_ MIMESAVETYPE* lpmst,
		_Out_opt_ ULONG* lpulWrapLines,
		_Out_opt_ bool* pDoAdrBook);
	_Check_return_ HRESULT GetConversionFromEMLOptions(
		_In_ CWnd* pParentWnd,
		_Out_opt_ CCSFLAGS* lpConvertFlags,
		_Out_opt_ bool* pDoAdrBook,
		_Out_opt_ bool* pDoApply,
		_Out_opt_ HCHARSET* phCharSet,
		_Out_opt_ CSETAPPLYTYPE* pcSetApplyType,
		_Out_opt_ bool* pbUnicode);

	void displayError(const std::wstring& errString);
	void OutputToDbgView(const std::wstring& szMsg);

	std::wstring PromptServerName();
} // namespace ui::mapiui