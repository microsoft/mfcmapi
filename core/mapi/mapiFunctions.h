// Stand alone MAPI functions
#pragma once

namespace mapi
{
	// Safely cast across MAPI interfaces. Result is addrefed and must be released.
	template <class T> T safe_cast(IUnknown* src);

	_Check_return_ ULONG GetMAPIObjectType(_In_opt_ LPMAPIPROP lpMAPIProp);

	_Check_return_ std::wstring EncodeID(ULONG cbEID, _In_ LPENTRYID rgbID);
	_Check_return_ std::wstring DecodeID(ULONG cbBuffer, _In_count_(cbBuffer) LPBYTE lpbBuffer);

	_Check_return_ HRESULT GetPropsNULL(
		_In_ LPMAPIPROP lpMAPIProp,
		ULONG ulFlags,
		_Out_ ULONG* lpcValues,
		_Deref_out_opt_ LPSPropValue* lppPropArray);

	_Check_return_ LPSPropValue GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);
	_Check_return_ LPSPropValue GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);
} // namespace mapi