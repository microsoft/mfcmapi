// Stand alone MAPI functions
#pragma once

namespace mapi
{
	// Safely cast across MAPI interfaces. Result is addrefed and must be released.
	template <class T> T safe_cast(IUnknown* src);

	_Check_return_ ULONG GetMAPIObjectType(_In_opt_ LPMAPIPROP lpMAPIProp);

	_Check_return_ std::wstring EncodeID(ULONG cbEID, _In_ LPENTRYID rgbID);
	_Check_return_ std::wstring DecodeID(ULONG cbBuffer, _In_count_(cbBuffer) LPBYTE lpbBuffer);
} // namespace mapi