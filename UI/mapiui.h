#pragma once
#include <Interpret/ExtraPropTags.h>

namespace ui
{
	namespace mapiui
	{
		void initCallbacks();

		LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB);

		void ExportMessages(_In_ LPMAPIFOLDER lpFolder, HWND hWnd);
		_Check_return_ HRESULT WriteAttachmentToFile(_In_ LPATTACH lpAttach, HWND hWnd);

		_Check_return_ HRESULT GetConversionToEMLOptions(
			_In_ CWnd* pParentWnd,
			_Out_ CCSFLAGS* lpConvertFlags,
			_Out_ const ENCODINGTYPE* lpet,
			_Out_ MIMESAVETYPE* lpmst,
			_Out_ ULONG* lpulWrapLines,
			_Out_ bool* pDoAdrBook);
		_Check_return_ HRESULT GetConversionFromEMLOptions(
			_In_ CWnd* pParentWnd,
			_Out_ CCSFLAGS* lpConvertFlags,
			_Out_ bool* pDoAdrBook,
			_Out_ bool* pDoApply,
			_Out_ HCHARSET* phCharSet,
			_Out_ CSETAPPLYTYPE* pcSetApplyType,
			_Out_opt_ bool* pbUnicode);
	} // namespace mapiui
} // namespace ui