// Stand alone MAPI functions
#pragma once

namespace mapi
{
	struct CopyDetails
	{
		bool valid{};
		ULONG flags{};
		GUID guid{};
		LPMAPIPROGRESS progress{};
		ULONG_PTR uiParam{};
		LPSPropTagArray excludedTags{};
		bool allocated{};
		void clean() const
		{
			if (progress) progress->Release();
			if (allocated) MAPIFreeBuffer(excludedTags);
		}
	};

	extern std::function<CopyDetails(
		HWND hWnd,
		_In_ LPMAPIPROP lpSource,
		LPCGUID lpGUID,
		_In_opt_ LPSPropTagArray lpTagArray,
		bool bIsAB)>
		GetCopyDetails;

	// Safely cast across MAPI interfaces. Result is addrefed and must be released.
	template <class T> T safe_cast(IUnknown* src);

	LPUNKNOWN CallOpenEntry(
		_In_opt_ LPMDB lpMDB,
		_In_opt_ LPADRBOOK lpAB,
		_In_opt_ LPMAPICONTAINER lpContainer,
		_In_opt_ LPMAPISESSION lpMAPISession,
		ULONG cbEntryID,
		_In_opt_ LPENTRYID lpEntryID,
		_In_opt_ LPCIID lpInterface,
		ULONG ulFlags,
		_Out_opt_ ULONG* ulObjTypeRet); // optional - can be NULL

	template <class T>
	T CallOpenEntry(
		_In_opt_ LPMDB lpMDB,
		_In_opt_ LPADRBOOK lpAB,
		_In_opt_ LPMAPICONTAINER lpContainer,
		_In_opt_ LPMAPISESSION lpMAPISession,
		ULONG cbEntryID,
		_In_opt_ LPENTRYID lpEntryID,
		_In_opt_ LPCIID lpInterface,
		ULONG ulFlags,
		_Out_opt_ ULONG* ulObjTypeRet) // optional - can be NULL
	{
		auto lpUnk = CallOpenEntry(
			lpMDB, lpAB, lpContainer, lpMAPISession, cbEntryID, lpEntryID, lpInterface, ulFlags, ulObjTypeRet);
		auto retVal = mapi::safe_cast<T>(lpUnk);
		if (lpUnk) lpUnk->Release();
		return retVal;
	}

	template <class T>
	T CallOpenEntry(
		_In_opt_ LPMDB lpMDB,
		_In_opt_ LPADRBOOK lpAB,
		_In_opt_ LPMAPICONTAINER lpContainer,
		_In_opt_ LPMAPISESSION lpMAPISession,
		_In_opt_ const SBinary* lpSBinary,
		_In_opt_ LPCIID lpInterface,
		ULONG ulFlags,
		_Out_opt_ ULONG* ulObjTypeRet) // optional - can be NULL
	{
		return CallOpenEntry<T>(
			lpMDB,
			lpAB,
			lpContainer,
			lpMAPISession,
			lpSBinary ? lpSBinary->cb : 0,
			reinterpret_cast<LPENTRYID>(lpSBinary ? lpSBinary->lpb : nullptr),
			lpInterface,
			ulFlags,
			ulObjTypeRet);
	}

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

	HRESULT CopyTo(
		HWND hWnd,
		_In_ LPMAPIPROP lpSource,
		_In_ LPMAPIPROP lpDest,
		LPCGUID lpGUID,
		_In_opt_ LPSPropTagArray lpTagArray,
		bool bIsAB,
		bool bAllowUI);

	_Check_return_ SBinary CopySBinary(_In_ const _SBinary& src, _In_ LPVOID parent = nullptr);
	_Check_return_ LPSBinary CopySBinary(_In_ const _SBinary* src);
} // namespace mapi