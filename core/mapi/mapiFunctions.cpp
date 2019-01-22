// Collection of useful MAPI functions
#include <core/stdafx.h>
#include <core/mapi/mapiFunctions.h>
#include <core/interpret/guid.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <core/utility/strings.h>
#include <core/utility/registry.h>
#include <core/mapi/mapiMemory.h>

namespace mapi
{
	// Safely cast across MAPI interfaces. Result is addrefed and must be released.
	template <class T> T safe_cast(IUnknown* src)
	{
		if (!src) return nullptr;

		auto iid = IID();
		// clang-format off
		if (std::is_same_v<T, LPUNKNOWN>) iid = IID_IUnknown;
		else if (std::is_same_v<T, LPMAPIFOLDER>) iid = IID_IMAPIFolder;
		else if (std::is_same_v<T, LPMAPICONTAINER>) iid = IID_IMAPIContainer;
		else if (std::is_same_v<T, LPMAILUSER>) iid = IID_IMailUser;
		else if (std::is_same_v<T, LPABCONT>) iid = IID_IABContainer;
		else if (std::is_same_v<T, LPMESSAGE>) iid = IID_IMessage;
		else if (std::is_same_v<T, LPMDB>) iid = IID_IMsgStore;
		else if (std::is_same_v<T, LPMAPIFORMINFO>) iid = IID_IMAPIFormInfo;
		else if (std::is_same_v<T, LPMAPIPROP>) iid = IID_IMAPIProp;
		else if (std::is_same_v<T, LPMAPIFORM>) iid = IID_IMAPIForm;
		else if (std::is_same_v<T, LPPERSISTMESSAGE>) iid = IID_IPersistMessage;
		else if (std::is_same_v<T, IAttachmentSecurity*>) iid = guid::IID_IAttachmentSecurity;
		else if (std::is_same_v<T, LPSERVICEADMIN2>) iid = IID_IMsgServiceAdmin2;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTORE>) iid = IID_IExchangeManageStore;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTORE3>) iid = IID_IExchangeManageStore3;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTORE4>) iid = IID_IExchangeManageStore4;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTORE5>) iid = guid::IID_IExchangeManageStore5;
		else if (std::is_same_v<T, LPEXCHANGEMANAGESTOREEX>) iid = guid::IID_IExchangeManageStoreEx;
		else if (std::is_same_v<T, IProxyStoreObject*>) iid = guid::IID_IProxyStoreObject;
		else if (std::is_same_v<T, LPMAPICLIENTSHUTDOWN>) iid = IID_IMAPIClientShutdown;
		else if (std::is_same_v<T, LPPROFSECT>) iid = IID_IProfSect;
		else if (std::is_same_v<T, LPATTACH>) iid = IID_IAttachment;
		else assert(false);
		// clang-format on

		T ret = nullptr;
		WC_H_S(src->QueryInterface(iid, reinterpret_cast<LPVOID*>(&ret)));
		output::DebugPrint(
			DBGGeneric,
			L"safe_cast: iid =%ws, src = %p, ret = %p\n",
			guid::GUIDToStringAndName(&iid).c_str(),
			src,
			ret);

		return ret;
	}

	template LPUNKNOWN safe_cast<LPUNKNOWN>(IUnknown* src);
	template LPMAPIFOLDER safe_cast<LPMAPIFOLDER>(IUnknown* src);
	template LPMAPICONTAINER safe_cast<LPMAPICONTAINER>(IUnknown* src);
	template LPMAILUSER safe_cast<LPMAILUSER>(IUnknown* src);
	template LPABCONT safe_cast<LPABCONT>(IUnknown* src);
	template LPMESSAGE safe_cast<LPMESSAGE>(IUnknown* src);
	template LPMDB safe_cast<LPMDB>(IUnknown* src);
	template LPMAPIFORMINFO safe_cast<LPMAPIFORMINFO>(IUnknown* src);
	template LPMAPIPROP safe_cast<LPMAPIPROP>(IUnknown* src);
	template LPMAPIFORM safe_cast<LPMAPIFORM>(IUnknown* src);
	template LPPERSISTMESSAGE safe_cast<LPPERSISTMESSAGE>(IUnknown* src);
	template IAttachmentSecurity* safe_cast<IAttachmentSecurity*>(IUnknown* src);
	template LPSERVICEADMIN2 safe_cast<LPSERVICEADMIN2>(IUnknown* src);
	template LPEXCHANGEMANAGESTORE safe_cast<LPEXCHANGEMANAGESTORE>(IUnknown* src);
	template LPEXCHANGEMANAGESTORE3 safe_cast<LPEXCHANGEMANAGESTORE3>(IUnknown* src);
	template LPEXCHANGEMANAGESTORE4 safe_cast<LPEXCHANGEMANAGESTORE4>(IUnknown* src);
	template LPEXCHANGEMANAGESTORE5 safe_cast<LPEXCHANGEMANAGESTORE5>(IUnknown* src);
	template LPEXCHANGEMANAGESTOREEX safe_cast<LPEXCHANGEMANAGESTOREEX>(IUnknown* src);
	template IProxyStoreObject* safe_cast<IProxyStoreObject*>(IUnknown* src);
	template LPMAPICLIENTSHUTDOWN safe_cast<LPMAPICLIENTSHUTDOWN>(IUnknown* src);
	template LPPROFSECT safe_cast<LPPROFSECT>(IUnknown* src);
	template LPATTACH safe_cast<LPATTACH>(IUnknown* src);

	// See list of types (like MAPI_FOLDER) in mapidefs.h
	_Check_return_ ULONG GetMAPIObjectType(_In_opt_ LPMAPIPROP lpMAPIProp)
	{
		ULONG ulObjType = 0;
		LPSPropValue lpProp = nullptr;

		if (!lpMAPIProp) return 0; // 0's not a valid Object type

		WC_MAPI_S(HrGetOneProp(lpMAPIProp, PR_OBJECT_TYPE, &lpProp));

		if (lpProp) ulObjType = lpProp->Value.ul;

		MAPIFreeBuffer(lpProp);
		return ulObjType;
	}

	const WORD kwBaseOffset = 0xAC00; // Hangul char range (AC00-D7AF)
	// Allocates with new, free with delete[]
	_Check_return_ std::wstring EncodeID(ULONG cbEID, _In_ LPENTRYID rgbID)
	{
		auto pbSrc = reinterpret_cast<LPBYTE>(rgbID);
		std::wstring wzIDEncoded;

		// rgbID is the item Entry ID or the attachment ID
		// cbID is the size in bytes of rgbID
		for (ULONG i = 0; i < cbEID; i++, pbSrc++)
		{
			wzIDEncoded += static_cast<WCHAR>(*pbSrc + kwBaseOffset);
		}

		// pwzIDEncoded now contains the entry ID encoded.
		return wzIDEncoded;
	}

	_Check_return_ std::wstring DecodeID(ULONG cbBuffer, _In_count_(cbBuffer) LPBYTE lpbBuffer)
	{
		if (cbBuffer % 2) return strings::emptystring;

		const auto cbDecodedBuffer = cbBuffer / 2;
		// Allocate memory for lpDecoded
		const auto lpDecoded = new BYTE[cbDecodedBuffer];
		if (!lpDecoded) return strings::emptystring;

		// Subtract kwBaseOffset from every character and place result in lpDecoded
		auto lpwzSrc = reinterpret_cast<LPWSTR>(lpbBuffer);
		auto lpDst = lpDecoded;
		for (ULONG i = 0; i < cbDecodedBuffer; i++, lpwzSrc++, lpDst++)
		{
			*lpDst = static_cast<BYTE>(*lpwzSrc - kwBaseOffset);
		}

		auto szBin = strings::BinToHexString(lpDecoded, cbDecodedBuffer, true);
		delete[] lpDecoded;
		return szBin;
	}

	_Check_return_ HRESULT GetPropsNULL(
		_In_ LPMAPIPROP lpMAPIProp,
		ULONG ulFlags,
		_Out_ ULONG* lpcValues,
		_Deref_out_opt_ LPSPropValue* lppPropArray)
	{
		auto hRes = S_OK;
		*lpcValues = NULL;
		*lppPropArray = nullptr;

		if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;
		LPSPropTagArray lpTags = nullptr;
		if (registry::useGetPropList)
		{
			output::DebugPrint(DBGGeneric, L"GetPropsNULL: Calling GetPropList\n");
			hRes = WC_MAPI(lpMAPIProp->GetPropList(ulFlags, &lpTags));

			if (hRes == MAPI_E_BAD_CHARWIDTH)
			{
				EC_MAPI_S(lpMAPIProp->GetPropList(NULL, &lpTags));
			}
			else
			{
				CHECKHRESMSG(hRes, IDS_NOPROPLIST);
			}
		}
		else
		{
			output::DebugPrint(DBGGeneric, L"GetPropsNULL: Calling GetProps(NULL) on %p\n", lpMAPIProp);
		}

		hRes = WC_H_GETPROPS(lpMAPIProp->GetProps(lpTags, ulFlags, lpcValues, lppPropArray));
		MAPIFreeBuffer(lpTags);

		return hRes;
	}

	// Returns LPSPropValue with value of a property
	// Uses GetProps and falls back to OpenProperty if the value is large
	// Free with MAPIFreeBuffer
	_Check_return_ LPSPropValue GetLargeProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
	{
		if (!lpMAPIProp) return nullptr;
		output::DebugPrint(DBGGeneric, L"GetLargeProp getting buffer from 0x%08X\n", ulPropTag);

		ULONG cValues = 0;
		LPSPropValue lpPropArray = nullptr;
		auto bSuccess = false;

		auto tag = SPropTagArray{1, {ulPropTag}};
		WC_H_GETPROPS_S(lpMAPIProp->GetProps(&tag, 0, &cValues, &lpPropArray));

		if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag) &&
			lpPropArray->Value.err == MAPI_E_NOT_ENOUGH_MEMORY)
		{
			output::DebugPrint(DBGGeneric, L"GetLargeProp property reported in GetProps as large.\n");
			MAPIFreeBuffer(lpPropArray);
			lpPropArray = nullptr;
			// need to get the data as a stream
			LPSTREAM lpStream = nullptr;

			WC_MAPI_S(lpMAPIProp->OpenProperty(
				ulPropTag, &IID_IStream, STGM_READ, 0, reinterpret_cast<LPUNKNOWN*>(&lpStream)));
			if (lpStream)
			{
				STATSTG StatInfo = {};
				lpStream->Stat(&StatInfo, STATFLAG_NONAME); // find out how much space we need

				// We're not going to try to support MASSIVE properties.
				if (!StatInfo.cbSize.HighPart)
				{
					lpPropArray = mapi::allocate<LPSPropValue>(sizeof(SPropValue));
					if (lpPropArray)
					{
						lpPropArray->ulPropTag = ulPropTag;

						if (StatInfo.cbSize.LowPart)
						{
							LPBYTE lpBuffer = nullptr;
							const auto ulBufferSize = StatInfo.cbSize.LowPart;
							ULONG ulTrailingNullSize = 0;
							switch (PROP_TYPE(ulPropTag))
							{
							case PT_STRING8:
								ulTrailingNullSize = sizeof(char);
								break;
							case PT_UNICODE:
								ulTrailingNullSize = sizeof(WCHAR);
								break;
							case PT_BINARY:
								break;
							default:
								break;
							}

							lpBuffer = mapi::allocate<LPBYTE>(
								static_cast<ULONG>(ulBufferSize + ulTrailingNullSize), lpPropArray);
							if (lpBuffer)
							{
								ULONG ulSizeRead = 0;
								const auto hRes = EC_MAPI(lpStream->Read(lpBuffer, ulBufferSize, &ulSizeRead));
								if (SUCCEEDED(hRes) && ulSizeRead == ulBufferSize)
								{
									switch (PROP_TYPE(ulPropTag))
									{
									case PT_STRING8:
										lpPropArray->Value.lpszA = reinterpret_cast<LPSTR>(lpBuffer);
										break;
									case PT_UNICODE:
										lpPropArray->Value.lpszW = reinterpret_cast<LPWSTR>(lpBuffer);
										break;
									case PT_BINARY:
										lpPropArray->Value.bin.cb = ulBufferSize;
										lpPropArray->Value.bin.lpb = lpBuffer;
										break;
									default:
										break;
									}

									bSuccess = true;
								}
							}
						}
						else
							bSuccess = true; // if LowPart was NULL, we return the empty buffer
					}
				}
			}
			if (lpStream) lpStream->Release();
		}
		else if (lpPropArray && cValues == 1 && lpPropArray->ulPropTag == ulPropTag)
		{
			output::DebugPrint(DBGGeneric, L"GetLargeProp GetProps found property.\n");
			bSuccess = true;
		}
		else if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag))
		{
			output::DebugPrint(
				DBGGeneric, L"GetLargeProp GetProps reported property as error 0x%08X.\n", lpPropArray->Value.err);
		}

		if (!bSuccess)
		{
			MAPIFreeBuffer(lpPropArray);
			lpPropArray = nullptr;
		}

		return lpPropArray;
	}

	// Returns LPSPropValue with value of a binary property
	// Free with MAPIFreeBuffer
	_Check_return_ LPSPropValue GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
	{
		return GetLargeProp(lpMAPIProp, CHANGE_PROP_TYPE(ulPropTag, PT_BINARY));
	}

	// Returns LPSPropValue with value of a string property
	// Free with MAPIFreeBuffer
	_Check_return_ LPSPropValue GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
	{
		return GetLargeProp(lpMAPIProp, CHANGE_PROP_TYPE(ulPropTag, PT_TSTRING));
	}
} // namespace mapi