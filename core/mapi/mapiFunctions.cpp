// Collection of useful MAPI functions
#include <core/stdafx.h>
#include <core/mapi/mapiFunctions.h>
#include <core/interpret/guid.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <core/utility/strings.h>
#include <core/utility/registry.h>
#include <core/mapi/mapiMemory.h>
#include <core/interpret/flags.h>
#include <core/mapi/mapiABFunctions.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/mapiProgress.h>
#include <core/mapi/cache/namedProps.h>
#include <core/mapi/mapiOutput.h>

namespace mapi
{
	// Safely cast across MAPI interfaces. Result is addrefed and must be released.
	// Adding interfaces:
	// 1 Add else if to safe_cast
	// 2 Add template to list to link an instance of the desired template
	// 3 Add USES_ statement to guid.cpp to ensure guid symbols is linked
	template <class T> T safe_cast(IUnknown* src)
	{
		if (!src) return nullptr;

		auto iid = IID();
		// clang-format off
		if (std::is_same_v<T, LPUNKNOWN>) iid = IID_IUnknown;
		else if (std::is_same_v<T, LPMAPISESSION>) iid = IID_IMAPISession;
		else if (std::is_same_v<T, LPMAPIFOLDER>) iid = IID_IMAPIFolder;
		else if (std::is_same_v<T, LPMAPICONTAINER>) iid = IID_IMAPIContainer;
		else if (std::is_same_v<T, LPMAILUSER>) iid = IID_IMailUser;
		else if (std::is_same_v<T, LPABCONT>) iid = IID_IABContainer;
		else if (std::is_same_v<T, LPMESSAGE>) iid = IID_IMessage;
		else if (std::is_same_v<T, LPMDB>) iid = IID_IMsgStore;
		else if (std::is_same_v<T, LPMAPIFORMINFO>) iid = IID_IMAPIFormInfo;
		else if (std::is_same_v<T, LPMAPIPROP>) iid = IID_IMAPIProp;
		else if (std::is_same_v<T, LPMAPIFORM>) iid = IID_IMAPIForm;
		else if (std::is_same_v<T, LPPERSISTMESSAGE>) iid =		IID_IPersistMessage;
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
		else if (std::is_same_v<T, LPOLKACCOUNT>) iid = IID_IOlkAccount;
		else if (std::is_same_v<T, LPMSCAPABILITIES>) iid = IID_IMSCapabilities;
		else assert(false);
		// clang-format on

		T ret = nullptr;
		WC_H_S(src->QueryInterface(iid, reinterpret_cast<LPVOID*>(&ret)));
		output::DebugPrint(
			output::dbgLevel::Generic,
			L"safe_cast: iid =%ws, src = %p, ret = %p\n",
			guid::GUIDToStringAndName(&iid).c_str(),
			src,
			ret);

		return ret;
	}

	template LPUNKNOWN safe_cast<LPUNKNOWN>(IUnknown* src);
	template LPMAPISESSION safe_cast<LPMAPISESSION>(IUnknown* src);
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
	template LPOLKACCOUNT safe_cast<LPOLKACCOUNT>(IUnknown* src);
	template LPMSCAPABILITIES safe_cast<LPMSCAPABILITIES>(IUnknown* src);

	LPUNKNOWN CallOpenEntry(
		_In_opt_ LPMDB lpMDB,
		_In_opt_ LPADRBOOK lpAB,
		_In_opt_ LPMAPICONTAINER lpContainer,
		_In_opt_ LPMAPISESSION lpMAPISession,
		const ULONG cbEntryID,
		const _In_opt_ LPENTRYID lpEntryID,
		_In_opt_ LPCIID lpInterface,
		ULONG ulFlags,
		_Out_opt_ ULONG* ulObjTypeRet) // optional - can be NULL
	{
		auto hRes = S_OK;
		ULONG ulObjType = 0;
		LPUNKNOWN lpUnk = nullptr;
		ULONG ulNoCacheFlags = 0;

		if (registry::forceMapiNoCache)
		{
			ulFlags |= MAPI_NO_CACHE;
		}

		// in case we need to retry without MAPI_NO_CACHE - do not add MAPI_NO_CACHE to ulFlags after this point
		if (MAPI_NO_CACHE & ulFlags) ulNoCacheFlags = ulFlags & ~MAPI_NO_CACHE;

		if (lpInterface && fIsSet(output::dbgLevel::Generic))
		{
			auto szGuid = guid::GUIDToStringAndName(lpInterface);
			output::DebugPrint(output::dbgLevel::Generic, L"CallOpenEntry: OpenEntry asking for %ws\n", szGuid.c_str());
		}

		if (lpMDB)
		{
			output::DebugPrint(
				output::dbgLevel::Generic, L"CallOpenEntry: Calling OpenEntry on MDB with ulFlags = 0x%X\n", ulFlags);
			hRes = WC_MAPI(lpMDB->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulFlags, &ulObjType, &lpUnk));
			if (hRes == MAPI_E_UNKNOWN_FLAGS && ulNoCacheFlags)
			{
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"CallOpenEntry 2nd attempt: Calling OpenEntry on MDB with ulFlags = 0x%X\n",
					ulNoCacheFlags);
				if (lpUnk) lpUnk->Release();
				lpUnk = nullptr;
				hRes = WC_MAPI(lpMDB->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulNoCacheFlags, &ulObjType, &lpUnk));
			}

			if (FAILED(hRes))
			{
				if (lpUnk) lpUnk->Release();
				lpUnk = nullptr;
			}
		}

		if (lpAB && !lpUnk)
		{
			output::DebugPrint(
				output::dbgLevel::Generic, L"CallOpenEntry: Calling OpenEntry on AB with ulFlags = 0x%X\n", ulFlags);
			hRes = WC_MAPI(lpAB->OpenEntry(
				cbEntryID,
				lpEntryID,
				nullptr, // no interface
				ulFlags,
				&ulObjType,
				&lpUnk));
			if (hRes == MAPI_E_UNKNOWN_FLAGS && ulNoCacheFlags)
			{
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"CallOpenEntry 2nd attempt: Calling OpenEntry on AB with ulFlags = 0x%X\n",
					ulNoCacheFlags);

				if (lpUnk) lpUnk->Release();
				lpUnk = nullptr;
				hRes = WC_MAPI(lpAB->OpenEntry(
					cbEntryID,
					lpEntryID,
					nullptr, // no interface
					ulNoCacheFlags,
					&ulObjType,
					&lpUnk));
			}

			if (FAILED(hRes))
			{
				if (lpUnk) lpUnk->Release();
				lpUnk = nullptr;
			}
		}

		if (lpContainer && !lpUnk)
		{
			output::DebugPrint(
				output::dbgLevel::Generic,
				L"CallOpenEntry: Calling OpenEntry on Container with ulFlags = 0x%X\n",
				ulFlags);
			hRes = WC_MAPI(lpContainer->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulFlags, &ulObjType, &lpUnk));
			if (hRes == MAPI_E_UNKNOWN_FLAGS && ulNoCacheFlags)
			{
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"CallOpenEntry 2nd attempt: Calling OpenEntry on Container with ulFlags = 0x%X\n",
					ulNoCacheFlags);

				if (lpUnk) lpUnk->Release();
				lpUnk = nullptr;
				hRes = WC_MAPI(
					lpContainer->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulNoCacheFlags, &ulObjType, &lpUnk));
			}

			if (FAILED(hRes))
			{
				if (lpUnk) lpUnk->Release();
				lpUnk = nullptr;
			}
		}

		if (lpMAPISession && !lpUnk)
		{
			output::DebugPrint(
				output::dbgLevel::Generic,
				L"CallOpenEntry: Calling OpenEntry on Session with ulFlags = 0x%X\n",
				ulFlags);
			hRes = WC_MAPI(lpMAPISession->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulFlags, &ulObjType, &lpUnk));
			if (hRes == MAPI_E_UNKNOWN_FLAGS && ulNoCacheFlags)
			{
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"CallOpenEntry 2nd attempt: Calling OpenEntry on Session with ulFlags = 0x%X\n",
					ulNoCacheFlags);

				if (lpUnk) lpUnk->Release();
				lpUnk = nullptr;
				hRes = WC_MAPI(
					lpMAPISession->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulNoCacheFlags, &ulObjType, &lpUnk));
			}

			if (FAILED(hRes))
			{
				if (lpUnk) lpUnk->Release();
				lpUnk = nullptr;
			}
		}

		if (lpUnk)
		{
			auto szFlags = flags::InterpretFlags(PROP_ID(PR_OBJECT_TYPE), static_cast<LONG>(ulObjType));
			output::DebugPrint(
				output::dbgLevel::Generic,
				L"OnOpenEntryID: Got object of type 0x%08X = %ws\n",
				ulObjType,
				szFlags.c_str());
		}

		if (ulObjTypeRet) *ulObjTypeRet = ulObjType;

		return lpUnk;
	}

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
			output::DebugPrint(output::dbgLevel::Generic, L"GetPropsNULL: Calling GetPropList\n");
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
			output::DebugPrint(output::dbgLevel::Generic, L"GetPropsNULL: Calling GetProps(NULL) on %p\n", lpMAPIProp);
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
		output::DebugPrint(output::dbgLevel::Generic, L"GetLargeProp getting buffer from 0x%08X\n", ulPropTag);

		ULONG cValues = 0;
		LPSPropValue lpPropArray = nullptr;
		auto bSuccess = false;

		auto tag = SPropTagArray{1, {ulPropTag}};
		WC_H_GETPROPS_S(lpMAPIProp->GetProps(&tag, 0, &cValues, &lpPropArray));

		if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag) &&
			lpPropArray->Value.err == MAPI_E_NOT_ENOUGH_MEMORY)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"GetLargeProp property reported in GetProps as large.\n");
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
										lpPropArray->Value.lpszW = strings::LPCBYTEToLPWSTR(lpBuffer);
										break;
									case PT_BINARY:
										mapi::setBin(lpPropArray) = {ulBufferSize, lpBuffer};
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
			output::DebugPrint(output::dbgLevel::Generic, L"GetLargeProp GetProps found property.\n");
			bSuccess = true;
		}
		else if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag))
		{
			output::DebugPrint(
				output::dbgLevel::Generic,
				L"GetLargeProp GetProps reported property as error 0x%08X.\n",
				lpPropArray->Value.err);
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

	std::function<CopyDetails(
		HWND hWnd,
		_In_ LPMAPIPROP lpSource,
		LPCGUID lpGUID,
		_In_opt_ LPSPropTagArray lpTagArray,
		bool bIsAB)>
		GetCopyDetails;

	// Performs CopyTo operation from source to destination, optionally prompting for exclusions
	// Does not save changes - caller should do this.
	HRESULT CopyTo(
		HWND hWnd,
		_In_ LPMAPIPROP lpSource,
		_In_ LPMAPIPROP lpDest,
		LPCGUID lpGUID,
		_In_opt_ LPSPropTagArray lpTagArray,
		bool bIsAB,
		bool bAllowUI)
	{
		if (!lpSource || !lpDest) return MAPI_E_INVALID_PARAMETER;

		auto copyDetails = GetCopyDetails && bAllowUI ? GetCopyDetails(hWnd, lpSource, lpGUID, lpTagArray, bIsAB)
													  : CopyDetails{true, 0, *lpGUID, nullptr, NULL, lpTagArray, false};
		if (copyDetails.valid)
		{
			auto lpProblems = LPSPropProblemArray{};
			const auto hRes = WC_MAPI(lpSource->CopyTo(
				0,
				nullptr,
				copyDetails.excludedTags,
				copyDetails.uiParam,
				copyDetails.progress,
				&copyDetails.guid,
				lpDest,
				copyDetails.flags,
				&lpProblems));
			if (lpProblems)
			{
				WC_PROBLEMARRAY(lpProblems);
				MAPIFreeBuffer(lpProblems);
			}

			copyDetails.clean();
			return hRes;
		}

		return MAPI_E_USER_CANCEL;
	}

	_Check_return_ SBinary CopySBinary(_In_ const _SBinary& src, _In_opt_ const VOID* parent)
	{
		const auto dst = SBinary{src.cb, mapi::allocate<LPBYTE>(src.cb, parent)};
		if (src.cb) CopyMemory(dst.lpb, src.lpb, src.cb);
		return dst;
	}

	_Check_return_ LPSBinary CopySBinary(_In_ const _SBinary* src)
	{
		if (!src) return nullptr;
		const auto dst = mapi::allocate<LPSBinary>(ULONG{sizeof(SBinary)});
		if (dst)
		{
			*dst = CopySBinary(*src, dst);
		}

		return dst;
	}

	///////////////////////////////////////////////////////////////////////////////
	// CopyString()
	//
	// Parameters
	// lpszDestination - Address of pointer to destination string
	// szSource - Pointer to source string
	// parent - Pointer to parent object (not, however, pointer to pointer!)
	//
	// Purpose
	// Uses MAPI to allocate a new string (szDestination) and copy szSource into it
	// Uses parent as the parent for MAPIAllocateMore if possible
	_Check_return_ LPSTR CopyStringA(_In_z_ LPCSTR src, _In_opt_ const VOID* pParent)
	{
		if (!src) return nullptr;
		auto cb = strnlen_s(src, RSIZE_MAX) + 1;
		auto dst = mapi::allocate<LPSTR>(static_cast<ULONG>(cb), pParent);

		if (dst)
		{
			EC_W32_S(strcpy_s(dst, cb, src));
		}

		return dst;
	}

	_Check_return_ LPWSTR CopyStringW(_In_z_ LPCWSTR src, _In_opt_ const VOID* pParent)
	{
		if (!src) return nullptr;
		auto cch = wcsnlen_s(src, RSIZE_MAX) + 1;
		const auto cb = cch * sizeof(WCHAR);
		auto dst = mapi::allocate<LPWSTR>(static_cast<ULONG>(cb), pParent);

		if (dst)
		{
			EC_W32_S(wcscpy_s(dst, cch, src));
		}

		return dst;
	}

	// I don't use MAPIOID.h, which is needed to deal with PR_ATTACH_TAG, but if I did, here's how to include it
	/*
	#include <mapiguid.h>
	#define USES_OID_TNEF
	#define USES_OID_OLE
	#define USES_OID_OLE1
	#define USES_OID_OLE1_STORAGE
	#define USES_OID_OLE2
	#define USES_OID_OLE2_STORAGE
	#define USES_OID_MAC_BINARY
	#define USES_OID_MIMETAG
	#define INITOID
	// Major hack to get MAPIOID to compile
	#define _MAC
	#include <MAPIOID.h>
	#undef _MAC
	*/

	// Concatenate two property arrays without duplicates
	_Check_return_ LPSPropTagArray
	ConcatSPropTagArrays(_In_ LPSPropTagArray lpArray1, _In_opt_ LPSPropTagArray lpArray2)
	{
		auto hRes = S_OK;

		// Add the sizes of the passed in arrays (0 if they were NULL)
		auto iNewArraySize = lpArray1 ? lpArray1->cValues : 0;

		if (lpArray2 && lpArray1)
		{
			for (ULONG iSourceArray = 0; iSourceArray < lpArray2->cValues; iSourceArray++)
			{
				if (!IsDuplicateProp(lpArray1, getTag(lpArray2, iSourceArray)))
				{
					iNewArraySize++;
				}
			}
		}
		else
		{
			iNewArraySize = iNewArraySize + (lpArray2 ? lpArray2->cValues : 0);
		}

		if (!iNewArraySize) return nullptr;

		// Allocate memory for the new prop tag array
		auto lpLocalArray = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(iNewArraySize));
		if (lpLocalArray)
		{
			ULONG iTargetArray = 0;
			if (lpArray1)
			{
				for (ULONG iSourceArray = 0; iSourceArray < lpArray1->cValues; iSourceArray++)
				{
					if (PROP_TYPE(getTag(lpArray1, iSourceArray)) != PT_NULL) // ditch bad props
					{
						setTag(lpLocalArray, iTargetArray++) = getTag(lpArray1, iSourceArray);
					}
				}
			}

			if (lpArray2)
			{
				for (ULONG iSourceArray = 0; iSourceArray < lpArray2->cValues; iSourceArray++)
				{
					if (PROP_TYPE(getTag(lpArray2, iSourceArray)) != PT_NULL) // ditch bad props
					{
						if (!IsDuplicateProp(lpArray1, getTag(lpArray2, iSourceArray)))
						{
							setTag(lpLocalArray, iTargetArray++) = getTag(lpArray2, iSourceArray);
						}
					}
				}
			}

			// <= since we may have thrown some PT_NULL tags out - just make sure we didn't overrun.
			hRes = EC_H(iTargetArray <= iNewArraySize ? S_OK : MAPI_E_CALL_FAILED);

			// since we may have ditched some tags along the way, reset our size
			lpLocalArray->cValues = iTargetArray;

			if (FAILED(hRes))
			{
				MAPIFreeBuffer(lpLocalArray);
				lpLocalArray = nullptr;
			}
		}

		return lpLocalArray;
	}

	_Check_return_ SBinaryArray GetEntryIDs(_In_ LPMAPITABLE table)
	{
		if (!table) return {};
		auto tag = SPropTagArray{1, {PR_ENTRYID}};
		auto hRes = EC_MAPI(table->SetColumns(&tag, TBL_BATCH));
		if (FAILED(hRes)) return {};

		ULONG ulRowCount = 0;
		EC_MAPI_S(table->GetRowCount(0, &ulRowCount));
		if (!ulRowCount || ulRowCount >= ULONG_MAX / sizeof(SBinary)) return {};

		auto sbaEID = SBinaryArray{0, mapi::allocate<LPSBinary>(sizeof(SBinary) * ulRowCount)};
		if (sbaEID.lpbin)
		{
			auto pRow = LPSRowSet();
			ULONG ulRowsCopied;
			for (ulRowsCopied = 0; ulRowsCopied < ulRowCount; ulRowsCopied++)
			{
				if (pRow) FreeProws(pRow);
				pRow = nullptr;
				hRes = EC_MAPI(table->QueryRows(1, NULL, &pRow));
				if (FAILED(hRes) || !pRow || !pRow->cRows) break;

				if (pRow && PT_ERROR != PROP_TYPE(pRow->aRow->lpProps[0].ulPropTag))
				{
					sbaEID.lpbin[ulRowsCopied] = CopySBinary(mapi::getBin(pRow->aRow->lpProps[0]), sbaEID.lpbin);
				}
			}

			if (pRow) FreeProws(pRow);
			sbaEID.cValues = ulRowsCopied;
		}

		return sbaEID;
	}

	// Gets entry IDs of the messages in the source folder and copies to destination using a single call to CopyMessage
	void CopyMessagesBatch(
		_In_ LPMAPIFOLDER lpSrcFolder,
		_In_ LPMAPIFOLDER lpDestFolder,
		bool bCopyAssociatedContents,
		bool bMove,
		_In_ HWND hWnd)
	{
		if (!lpSrcFolder || !lpDestFolder) return;
		LPMAPITABLE lpSrcContents = nullptr;
		EC_MAPI_S(lpSrcFolder->GetContentsTable(
			fMapiUnicode | (bCopyAssociatedContents ? MAPI_ASSOCIATED : NULL), &lpSrcContents));
		if (!lpSrcContents) return;

		auto sbaEID = GetEntryIDs(lpSrcContents);
		auto lpProgress = mapiui::GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK
		auto ulCopyFlags = bMove ? MESSAGE_MOVE : 0;

		if (lpProgress) ulCopyFlags |= MESSAGE_DIALOG;

		EC_MAPI_S(lpSrcFolder->CopyMessages(
			&sbaEID,
			&IID_IMAPIFolder,
			lpDestFolder,
			lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
			lpProgress,
			ulCopyFlags));

		if (lpProgress) lpProgress->Release();
		MAPIFreeBuffer(sbaEID.lpbin);
		lpSrcContents->Release();
	}

	_Check_return_ HRESULT CopyMessage(
		_In_ const SBinary& bin,
		_In_ LPMAPIFOLDER lpSrcFolder,
		_In_ LPMAPIFOLDER lpDestFolder,
		bool bMove,
		_In_ HWND hWnd)
	{
		if (!bin.lpb || !lpSrcFolder || !lpDestFolder) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(output::dbgLevel::Generic, L"Source Message =\n");
		output::outputBinary(output::dbgLevel::Generic, nullptr, bin);

		auto sbaEID = SBinaryArray{1, const_cast<LPSBinary>(&bin)};

		auto lpProgress = mapiui::GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK

		auto ulCopyFlags = bMove ? MESSAGE_MOVE : 0;
		if (lpProgress) ulCopyFlags |= MESSAGE_DIALOG;

		const auto hRes = EC_MAPI(lpSrcFolder->CopyMessages(
			&sbaEID,
			&IID_IMAPIFolder,
			lpDestFolder,
			lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
			lpProgress,
			ulCopyFlags));

		if (hRes == S_OK) output::DebugPrint(output::dbgLevel::Generic, L"Message copied\n");

		if (lpProgress) lpProgress->Release();
		return hRes;
	}

	// Copies message from source folder to destination using multiple calls to CopyMessage
	// Gets entry IDs of the messages in the source folder and copies to destination using a single call to CopyMessage
	void CopyMessagesIterate(
		_In_ LPMAPIFOLDER lpSrcFolder,
		_In_ LPMAPIFOLDER lpDestFolder,
		bool bCopyAssociatedContents,
		bool bMove,
		_In_ HWND hWnd)
	{
		if (!lpSrcFolder || !lpDestFolder) return;
		LPMAPITABLE lpSrcContents = nullptr;
		auto hRes = EC_MAPI(lpSrcFolder->GetContentsTable(
			fMapiUnicode | (bCopyAssociatedContents ? MAPI_ASSOCIATED : NULL), &lpSrcContents));
		if (!lpSrcContents) return;

		auto ulRowCount = ULONG();
		auto tag = SPropTagArray{1, {PR_ENTRYID}};
		hRes = EC_MAPI(lpSrcContents->SetColumns(&tag, TBL_BATCH));
		if (SUCCEEDED(hRes))
		{
			EC_MAPI_S(lpSrcContents->GetRowCount(0, &ulRowCount));
		}

		if (ulRowCount)
		{
			auto pRow = LPSRowSet();
			for (ULONG ulRowsCopied = 0; ulRowsCopied < ulRowCount; ulRowsCopied++)
			{
				if (pRow) FreeProws(pRow);
				pRow = nullptr;
				hRes = EC_MAPI(lpSrcContents->QueryRows(1, NULL, &pRow));
				if (FAILED(hRes) || !pRow || !pRow->cRows) break;

				if (PROP_TYPE(pRow->aRow->lpProps[0].ulPropTag) != PT_ERROR)
				{
					hRes =
						WC_H(CopyMessage(mapi::getBin(pRow->aRow->lpProps[0]), lpSrcFolder, lpDestFolder, bMove, hWnd));
				}

				if (S_OK != hRes) output::DebugPrint(output::dbgLevel::Generic, L"Message Copy Failed\n");
			}

			if (pRow) FreeProws(pRow);
		}

		lpSrcContents->Release();
	}

	// May not behave correctly if lpSrcFolder == lpDestFolder
	// We can check that the pointers aren't equal, but they could be different
	// and still refer to the same folder.
	void CopyFolderContents(
		_In_ LPMAPIFOLDER lpSrcFolder,
		_In_ LPMAPIFOLDER lpDestFolder,
		bool bCopyAssociatedContents,
		bool bMove,
		bool bSingleCall,
		_In_ HWND hWnd)
	{
		output::DebugPrint(
			output::dbgLevel::Generic,
			L"CopyFolderContents: lpSrcFolder = %p, lpDestFolder = %p, bCopyAssociatedContents = %d, bMove = %d\n",
			lpSrcFolder,
			lpDestFolder,
			bCopyAssociatedContents,
			bMove);

		if (!lpSrcFolder || !lpDestFolder) return;

		if (bSingleCall)
		{
			CopyMessagesBatch(lpSrcFolder, lpDestFolder, bCopyAssociatedContents, bMove, hWnd);
		}
		else
		{
			CopyMessagesIterate(lpSrcFolder, lpDestFolder, bCopyAssociatedContents, bMove, hWnd);
		}
	}

	_Check_return_ HRESULT CopyFolderRules(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bReplace)
	{
		if (!lpSrcFolder || !lpDestFolder) return MAPI_E_INVALID_PARAMETER;
		LPEXCHANGEMODIFYTABLE lpSrcTbl = nullptr;
		LPEXCHANGEMODIFYTABLE lpDstTbl = nullptr;

		auto hRes = EC_MAPI(lpSrcFolder->OpenProperty(
			PR_RULES_TABLE,
			const_cast<LPGUID>(&IID_IExchangeModifyTable),
			0,
			MAPI_DEFERRED_ERRORS,
			reinterpret_cast<LPUNKNOWN*>(&lpSrcTbl)));

		if (SUCCEEDED(hRes))
		{
			hRes = EC_MAPI(lpDestFolder->OpenProperty(
				PR_RULES_TABLE,
				const_cast<LPGUID>(&IID_IExchangeModifyTable),
				0,
				MAPI_DEFERRED_ERRORS,
				reinterpret_cast<LPUNKNOWN*>(&lpDstTbl)));
		}

		if (lpSrcTbl && lpDstTbl)
		{
			LPMAPITABLE lpTable = nullptr;
			lpSrcTbl->GetTable(0, &lpTable);

			if (lpTable)
			{
				static const SizedSPropTagArray(9, ruleTags) = {
					9,
					{PR_RULE_ACTIONS,
					 PR_RULE_CONDITION,
					 PR_RULE_LEVEL,
					 PR_RULE_NAME,
					 PR_RULE_PROVIDER,
					 PR_RULE_PROVIDER_DATA,
					 PR_RULE_SEQUENCE,
					 PR_RULE_STATE,
					 PR_RULE_USER_FLAGS},
				};

				hRes = EC_MAPI(lpTable->SetColumns(LPSPropTagArray(&ruleTags), 0));

				LPSRowSet lpRows = nullptr;

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI(HrQueryAllRows(lpTable, nullptr, nullptr, nullptr, 0, &lpRows));
				}

				if (lpRows && lpRows->cRows < MAXNewROWLIST)
				{
					auto lpTempList = mapi::allocate<LPROWLIST>(CbNewROWLIST(lpRows->cRows));
					if (lpTempList)
					{
						lpTempList->cEntries = lpRows->cRows;

						for (ULONG iArrayPos = 0; iArrayPos < lpRows->cRows; iArrayPos++)
						{
							lpTempList->aEntries[iArrayPos].ulRowFlags = ROW_ADD;
							lpTempList->aEntries[iArrayPos].rgPropVals = mapi::allocate<LPSPropValue>(
								lpRows->aRow[iArrayPos].cValues * sizeof(SPropValue), lpTempList);
							if (lpTempList->aEntries[iArrayPos].rgPropVals)
							{
								ULONG ulDst = 0;
								for (ULONG ulSrc = 0; ulSrc < lpRows->aRow[iArrayPos].cValues; ulSrc++)
								{
									if (lpRows->aRow[iArrayPos].lpProps[ulSrc].ulPropTag == PR_RULE_PROVIDER_DATA)
									{
										const auto& bin = mapi::getBin(lpRows->aRow[iArrayPos].lpProps[ulSrc]);
										if (!bin.cb || !bin.lpb)
										{
											// PR_RULE_PROVIDER_DATA was NULL - we don't want this
											continue;
										}
									}

									hRes = EC_H(MyPropCopyMore(
										&lpTempList->aEntries[iArrayPos].rgPropVals[ulDst],
										&lpRows->aRow[iArrayPos].lpProps[ulSrc],
										MAPIAllocateMore,
										lpTempList));
									ulDst++;
								}

								lpTempList->aEntries[iArrayPos].cValues = ulDst;
							}
						}

						ULONG ulFlags = 0;
						if (bReplace) ulFlags = ROWLIST_REPLACE;

						if (SUCCEEDED(hRes))
						{
							hRes = EC_MAPI(lpDstTbl->ModifyTable(ulFlags, lpTempList));
						}

						MAPIFreeBuffer(lpTempList);
					}

					FreeProws(lpRows);
				}

				lpTable->Release();
			}
		}

		if (lpDstTbl) lpDstTbl->Release();
		if (lpSrcTbl) lpSrcTbl->Release();
		return hRes;
	}

	// Copy a property using the stream interface
	// Does not call SaveChanges
	_Check_return_ HRESULT CopyPropertyAsStream(
		_In_ LPMAPIPROP lpSourcePropObj,
		_In_ LPMAPIPROP lpTargetPropObj,
		ULONG ulSourceTag,
		ULONG ulTargetTag)
	{
		LPSTREAM lpStmSource = nullptr;
		LPSTREAM lpStmTarget = nullptr;
		LARGE_INTEGER li = {};
		ULARGE_INTEGER uli = {};
		ULARGE_INTEGER ulBytesRead = {};
		ULARGE_INTEGER ulBytesWritten = {};

		if (!lpSourcePropObj || !lpTargetPropObj || !ulSourceTag || !ulTargetTag) return MAPI_E_INVALID_PARAMETER;
		if (PROP_TYPE(ulSourceTag) != PROP_TYPE(ulTargetTag)) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(lpSourcePropObj->OpenProperty(
			ulSourceTag, &IID_IStream, STGM_READ | STGM_DIRECT, NULL, reinterpret_cast<LPUNKNOWN*>(&lpStmSource)));

		if (SUCCEEDED(hRes))
		{
			hRes = EC_MAPI(lpTargetPropObj->OpenProperty(
				ulTargetTag,
				&IID_IStream,
				STGM_READWRITE | STGM_DIRECT,
				MAPI_CREATE | MAPI_MODIFY,
				reinterpret_cast<LPUNKNOWN*>(&lpStmTarget)));
		}

		if (lpStmSource && lpStmTarget)
		{
			li.QuadPart = 0;
			uli.QuadPart = MAXLONGLONG;

			hRes = EC_MAPI(lpStmSource->Seek(li, STREAM_SEEK_SET, nullptr));

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpStmTarget->Seek(li, STREAM_SEEK_SET, nullptr));
			}

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpStmSource->CopyTo(lpStmTarget, uli, &ulBytesRead, &ulBytesWritten));
			}

			if (SUCCEEDED(hRes))
			{
				// This may not be necessary since we opened with STGM_DIRECT
				hRes = EC_MAPI(lpStmTarget->Commit(STGC_DEFAULT));
			}
		}

		if (lpStmTarget) lpStmTarget->Release();
		if (lpStmSource) lpStmSource->Release();
		return hRes;
	}

	// Allocates and creates a restriction that looks for existence of
	// a particular property that matches the given string
	// If parent is passed in, it is used as the allocation parent.
	_Check_return_ LPSRestriction CreatePropertyStringRestriction(
		ULONG ulPropTag,
		_In_ const std::wstring& szString,
		ULONG ulFuzzyLevel,
		_In_opt_ const VOID* parent)
	{
		if (PROP_TYPE(ulPropTag) != PT_UNICODE) return nullptr;
		if (szString.empty()) return nullptr;

		// Allocate and create our SRestriction
		// Allocate base memory:
		const auto lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), parent);
		if (lpRes)
		{
			const auto lpAllocationParent = parent ? parent : lpRes;

			// Root Node
			lpRes->rt = RES_AND;
			lpRes->res.resAnd.cRes = 2;
			const auto lpResLevel1 = mapi::allocate<LPSRestriction>(sizeof(SRestriction) * 2, lpAllocationParent);
			lpRes->res.resAnd.lpRes = lpResLevel1;
			if (lpResLevel1)
			{
				lpResLevel1[0].rt = RES_EXIST;
				lpResLevel1[0].res.resExist.ulPropTag = ulPropTag;
				lpResLevel1[0].res.resExist.ulReserved1 = 0;
				lpResLevel1[0].res.resExist.ulReserved2 = 0;

				lpResLevel1[1].rt = RES_CONTENT;
				lpResLevel1[1].res.resContent.ulPropTag = ulPropTag;
				lpResLevel1[1].res.resContent.ulFuzzyLevel = ulFuzzyLevel;
				const auto lpspvSubject = mapi::allocate<LPSPropValue>(sizeof(SPropValue), lpAllocationParent);
				lpResLevel1[1].res.resContent.lpProp = lpspvSubject;
				if (lpspvSubject)
				{
					// Allocate and fill out properties:
					lpspvSubject->ulPropTag = ulPropTag;
					lpspvSubject->Value.lpszW = CopyStringW(szString.c_str(), lpAllocationParent);
				}

				output::DebugPrint(output::dbgLevel::Generic, L"CreatePropertyStringRestriction built restriction:\n");
				output::outputRestriction(output::dbgLevel::Generic, nullptr, lpRes, nullptr);
			}
		}

		return lpRes;
	}

	_Check_return_ LPSRestriction
	CreateRangeRestriction(ULONG ulPropTag, _In_ const std::wstring& szString, _In_opt_ const VOID* parent)
	{
		if (szString.empty()) return nullptr;
		if (PROP_TYPE(ulPropTag) != PT_UNICODE) return nullptr;

		// Allocate and create our SRestriction
		// Allocate base memory:
		const auto lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), parent);

		if (lpRes)
		{
			const auto lpAllocationParent = parent ? parent : lpRes;

			// Root Node
			lpRes->rt = RES_AND;
			lpRes->res.resAnd.cRes = 2;
			const auto lpResLevel1 = mapi::allocate<LPSRestriction>(sizeof(SRestriction) * 2, lpAllocationParent);
			lpRes->res.resAnd.lpRes = lpResLevel1;

			if (lpResLevel1)
			{
				lpResLevel1[0].rt = RES_EXIST;
				lpResLevel1[0].res.resExist.ulPropTag = ulPropTag;
				lpResLevel1[0].res.resExist.ulReserved1 = 0;
				lpResLevel1[0].res.resExist.ulReserved2 = 0;

				lpResLevel1[1].rt = RES_PROPERTY;
				lpResLevel1[1].res.resProperty.ulPropTag = ulPropTag;
				lpResLevel1[1].res.resProperty.relop = RELOP_GE;
				const auto lpspvSubject = mapi::allocate<LPSPropValue>(sizeof(SPropValue), lpAllocationParent);
				lpResLevel1[1].res.resProperty.lpProp = lpspvSubject;

				if (lpspvSubject)
				{
					// Allocate and fill out properties:
					lpspvSubject->ulPropTag = ulPropTag;
					lpspvSubject->Value.lpszW = CopyStringW(szString.c_str(), lpAllocationParent);
				}
			}

			output::DebugPrint(output::dbgLevel::Generic, L"CreateRangeRestriction built restriction:\n");
			output::outputRestriction(output::dbgLevel::Generic, nullptr, lpRes, nullptr);
		}

		return lpRes;
	}

	_Check_return_ HRESULT DeleteProperty(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
	{
		LPSPropProblemArray pProbArray = nullptr;

		if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

		if (PROP_TYPE(ulPropTag) == PT_ERROR) ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED);

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"DeleteProperty: Deleting prop 0x%08X from MAPI item %p.\n",
			ulPropTag,
			lpMAPIProp);

		SPropTagArray ptaTag = {1, {ulPropTag}};
		auto hRes = EC_MAPI(lpMAPIProp->DeleteProps(&ptaTag, &pProbArray));

		if (hRes == S_OK && pProbArray)
		{
			WC_PROBLEMARRAY(pProbArray);
			if (pProbArray->cProblem == 1)
			{
				hRes = pProbArray->aProblem[0].scode;
			}
		}

		if (SUCCEEDED(hRes))
		{
			hRes = EC_MAPI(lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
		}

		MAPIFreeBuffer(pProbArray);

		return hRes;
	}

	// Delete items to the wastebasket of the passed in mdb, if it exists.
	_Check_return_ HRESULT
	DeleteToDeletedItems(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpSourceFolder, _In_ LPENTRYLIST lpEIDs, _In_ HWND hWnd)
	{
		if (!lpMDB || !lpSourceFolder || !lpEIDs) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"DeleteToDeletedItems: Deleting from folder %p in store %p\n",
			lpSourceFolder,
			lpMDB);

		auto hRes = S_OK;
		auto lpWasteFolder = OpenDefaultFolder(DEFAULT_DELETEDITEMS, lpMDB);
		if (lpWasteFolder)
		{
			auto lpProgress = mapiui::GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK

			auto ulCopyFlags = MESSAGE_MOVE;

			if (lpProgress) ulCopyFlags |= MESSAGE_DIALOG;

			hRes = EC_MAPI(lpSourceFolder->CopyMessages(
				lpEIDs,
				nullptr, // default interface
				lpWasteFolder,
				lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
				lpProgress,
				ulCopyFlags));

			if (lpProgress) lpProgress->Release();
		}

		if (lpWasteFolder) lpWasteFolder->Release();
		return hRes;
	}

	_Check_return_ bool
	FindPropInPropTagArray(_In_ LPSPropTagArray lpspTagArray, ULONG ulPropToFind, _Out_ ULONG* lpulRowFound) noexcept
	{
		*lpulRowFound = 0;
		if (!lpspTagArray) return false;
		for (ULONG i = 0; i < lpspTagArray->cValues; i++)
		{
			if (PROP_ID(ulPropToFind) == PROP_ID(getTag(lpspTagArray, i)))
			{
				*lpulRowFound = i;
				return true;
			}
		}

		return false;
	}

	_Check_return_ LPSBinary GetInboxEntryId(_In_ LPMDB lpMDB)
	{
		output::DebugPrint(output::dbgLevel::Generic, L"GetInboxEntryId: getting Inbox\n");

		if (!lpMDB) return {};

		auto receiveEid = SBinary{};
		EC_MAPI_S(lpMDB->GetReceiveFolder(
			const_cast<LPTSTR>(_T("IPM.Note")), // STRING_OK this is the class of message we want
			fMapiUnicode, // flags
			&receiveEid.cb, // size and...
			reinterpret_cast<LPENTRYID*>(&receiveEid.lpb), // value of entry ID
			nullptr)); // returns a message class if not NULL

		auto eid = LPSBinary{};
		if (receiveEid.cb && receiveEid.lpb)
		{
			eid = CopySBinary(&receiveEid);
		}

		MAPIFreeBuffer(receiveEid.lpb);
		return eid;
	}

	_Check_return_ LPMAPIFOLDER GetInbox(_In_ LPMDB lpMDB)
	{
		if (!lpMDB) return nullptr;

		output::DebugPrint(output::dbgLevel::Generic, L"GetInbox: getting Inbox from %p\n", lpMDB);

		const auto eid = GetInboxEntryId(lpMDB);

		LPMAPIFOLDER lpInbox = nullptr;
		if (eid)
		{
			// Get the Inbox...
			lpInbox =
				CallOpenEntry<LPMAPIFOLDER>(lpMDB, nullptr, nullptr, nullptr, eid, nullptr, MAPI_BEST_ACCESS, nullptr);
		}

		MAPIFreeBuffer(eid);
		return lpInbox;
	}

	_Check_return_ LPMAPIFOLDER GetParentFolder(_In_ LPMAPIFOLDER lpChildFolder, _In_ LPMDB lpMDB)
	{
		if (!lpChildFolder) return nullptr;
		ULONG cProps = {};
		LPSPropValue lpProps = nullptr;
		LPMAPIFOLDER lpParentFolder = nullptr;

		static auto tag = SPropTagArray{1, {PR_PARENT_ENTRYID}};

		// Get PR_PARENT_ENTRYID
		EC_H_GETPROPS_S(lpChildFolder->GetProps(&tag, fMapiUnicode, &cProps, &lpProps));
		if (lpProps && PT_ERROR != PROP_TYPE(lpProps[0].ulPropTag))
		{
			const auto& bin = mapi::getBin(lpProps[0]);
			lpParentFolder = CallOpenEntry<LPMAPIFOLDER>(
				lpMDB,
				nullptr,
				nullptr,
				nullptr,
				bin.cb,
				reinterpret_cast<LPENTRYID>(bin.lpb),
				nullptr,
				MAPI_BEST_ACCESS,
				nullptr);
		}

		MAPIFreeBuffer(lpProps);
		return lpParentFolder;
	}

	_Check_return_ LPSBinary GetSpecialFolderEID(_In_ LPMDB lpMDB, ULONG ulFolderPropTag)
	{
		if (!lpMDB) return {};

		output::DebugPrint(
			output::dbgLevel::Generic, L"GetSpecialFolderEID: getting 0x%X from %p\n", ulFolderPropTag, lpMDB);

		auto hRes = S_OK;
		LPSPropValue lpProp = nullptr;
		auto lpInbox = GetInbox(lpMDB);
		if (lpInbox)
		{
			hRes = WC_H_MSG(IDS_GETSPECIALFOLDERINBOXMISSINGPROP, HrGetOneProp(lpInbox, ulFolderPropTag, &lpProp));
			lpInbox->Release();
		}

		if (!lpProp)
		{
			// Open root container.
			auto lpRootFolder = CallOpenEntry<LPMAPIFOLDER>(
				lpMDB,
				nullptr,
				nullptr,
				nullptr,
				nullptr, // open root container
				nullptr,
				MAPI_BEST_ACCESS,
				nullptr);
			if (lpRootFolder)
			{
				hRes =
					EC_H_MSG(IDS_GETSPECIALFOLDERROOTMISSINGPROP, HrGetOneProp(lpRootFolder, ulFolderPropTag, &lpProp));
				lpRootFolder->Release();
			}
		}

		auto eid = LPSBinary{};
		if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag) && mapi::getBin(lpProp).cb)
		{
			eid = CopySBinary(&mapi::getBin(lpProp));
		}

		if (hRes == MAPI_E_NOT_FOUND)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"Special folder not found.\n");
		}

		MAPIFreeBuffer(lpProp);
		return eid;
	}

	_Check_return_ HRESULT
	IsAttachmentBlocked(_In_ LPMAPISESSION lpMAPISession, _In_z_ LPCWSTR pwszFileName, _Out_ bool* pfBlocked)
	{
		if (!lpMAPISession || !pwszFileName || !pfBlocked) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		auto bBlocked = BOOL{false};
		auto lpAttachSec = mapi::safe_cast<IAttachmentSecurity*>(lpMAPISession);
		if (lpAttachSec)
		{
			hRes = EC_MAPI(lpAttachSec->IsAttachmentBlocked(pwszFileName, &bBlocked));
			lpAttachSec->Release();
		}

		*pfBlocked = !!bBlocked;
		return hRes;
	}

	_Check_return_ bool IsDuplicateProp(_In_ LPSPropTagArray lpArray, ULONG ulPropTag) noexcept
	{
		if (!lpArray) return false;

		for (ULONG i = 0; i < lpArray->cValues; i++)
		{
			// They're dupes if the IDs are the same
			if (registry::allowDupeColumns)
			{
				if (getTag(lpArray, i) == ulPropTag) return true;
			}
			else
			{
				if (PROP_ID(getTag(lpArray, i)) == PROP_ID(ulPropTag)) return true;
			}
		}

		return false;
	}

	_Check_return_ HRESULT ManuallyEmptyFolder(_In_ LPMAPIFOLDER lpFolder, BOOL bAssoc, BOOL bHardDelete)
	{
		if (!lpFolder) return MAPI_E_INVALID_PARAMETER;

		LPSRowSet pRows = nullptr;
		ULONG iItemCount = 0;
		LPMAPITABLE lpContentsTable = nullptr;

		enum
		{
			eidPR_ENTRYID,
			eidNUM_COLS
		};

		static const SizedSPropTagArray(eidNUM_COLS, eidCols) = {eidNUM_COLS, {PR_ENTRYID}};

		// Get the table of contents of the folder
		auto hRes = WC_MAPI(lpFolder->GetContentsTable(bAssoc ? MAPI_ASSOCIATED : NULL, &lpContentsTable));

		if (SUCCEEDED(hRes) && lpContentsTable)
		{
			hRes = EC_MAPI(lpContentsTable->SetColumns(LPSPropTagArray(&eidCols), TBL_BATCH));

			// go to the first row
			if (SUCCEEDED(hRes))
			{
				EC_MAPI_S(lpContentsTable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));
			}

			// get rows and delete messages one at a time (slow, but might work when batch deletion fails)
			if (!FAILED(hRes))
			{
				for (;;)
				{
					if (pRows) FreeProws(pRows);
					pRows = nullptr;
					// Pull back a sizable block of rows to delete
					hRes = EC_MAPI(lpContentsTable->QueryRows(200, NULL, &pRows));
					if (FAILED(hRes) || !pRows || !pRows->cRows) break;

					for (ULONG iCurPropRow = 0; iCurPropRow < pRows->cRows; iCurPropRow++)
					{
						if (pRows->aRow[iCurPropRow].lpProps &&
							PR_ENTRYID == pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID].ulPropTag)
						{
							ENTRYLIST eid = {1, &mapi::setBin(pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID])};
							hRes = WC_MAPI(
								lpFolder->DeleteMessages(&eid, NULL, nullptr, bHardDelete ? DELETE_HARD_DELETE : NULL));
							if (SUCCEEDED(hRes)) iItemCount++;
						}
					}
				}
			}

			output::DebugPrint(output::dbgLevel::Generic, L"ManuallyEmptyFolder deleted %u items\n", iItemCount);
		}

		if (pRows) FreeProws(pRows);
		if (lpContentsTable) lpContentsTable->Release();
		return hRes;
	}

	// Converts vector<BYTE> to LPBYTE allocated with MAPIAllocateMore
	// Will only return nullptr on allocation failure. Even empty bin will return pointer to 0 so MAPI handles empty strings properly
	_Check_return_ LPBYTE ByteVectorToMAPI(const std::vector<BYTE>& bin, const VOID* parent)
	{
		const auto binsize = static_cast<ULONG>(bin.size() + sizeof(WCHAR));
		// We allocate a couple extra bytes (initialized to NULL) in case this buffer is printed.
		const auto lpBin = mapi::allocate<LPBYTE>(binsize, parent);
		if (lpBin)
		{
			memset(lpBin, 0, binsize);
			if (!bin.empty()) memcpy(lpBin, &bin[0], bin.size());
			return lpBin;
		}

		return nullptr;
	}

	const ULONG aulOneOffIDs[] = {
		dispidFormStorage,
		dispidPageDirStream,
		dispidFormPropStream,
		dispidScriptStream,
		dispidPropDefStream, // dispidPropDefStream must remain next to last in list
		dispidCustomFlag}; // dispidCustomFlag must remain last in list
	constexpr ULONG ulNumOneOffIDs = _countof(aulOneOffIDs);

	_Check_return_ HRESULT RemoveOneOff(_In_ LPMESSAGE lpMessage, bool bRemovePropDef)
	{
		if (!lpMessage) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(output::dbgLevel::NamedProp, L"RemoveOneOff - removing one off named properties.\n");

		auto hRes = S_OK;
		auto names = std::vector<MAPINAMEID>{};

		for (ULONG i = 0; i < ulNumOneOffIDs; i++)
		{
			auto name = MAPINAMEID{};
			name.lpguid = const_cast<LPGUID>(&guid::PSETID_Common);
			name.ulKind = MNID_ID;
			name.Kind.lID = aulOneOffIDs[i];
			names.emplace_back(name);
		}

		auto lpTags = cache::GetIDsFromNames(lpMessage, names, 0);
		if (lpTags)
		{
			LPSPropProblemArray lpProbArray = nullptr;

			output::DebugPrint(output::dbgLevel::NamedProp, L"RemoveOneOff - identified the following properties.\n");
			output::outputPropTagArray(output::dbgLevel::NamedProp, nullptr, lpTags);

			// The last prop is the flag value we'll be updating, don't count it
			lpTags->cValues = ulNumOneOffIDs - 1;

			// If we're not removing the prop def stream, then don't count it
			if (!bRemovePropDef)
			{
				lpTags->cValues = lpTags->cValues - 1;
			}

			hRes = EC_MAPI(lpMessage->DeleteProps(lpTags, &lpProbArray));
			if (SUCCEEDED(hRes))
			{
				if (lpProbArray)
				{
					output::DebugPrint(
						output::dbgLevel::NamedProp,
						L"RemoveOneOff - DeleteProps problem array:\n%ws\n",
						error::ProblemArrayToString(*lpProbArray).c_str());
				}

				ULONG cProp = 0;
				LPSPropValue lpCustomFlag = nullptr;

				// Grab dispidCustomFlag, the last tag in the array
				SPropTagArray pTag = {1, {CHANGE_PROP_TYPE(getTag(lpTags, ulNumOneOffIDs - 1), PT_LONG)}};

				hRes = WC_MAPI(lpMessage->GetProps(&pTag, fMapiUnicode, &cProp, &lpCustomFlag));
				if (SUCCEEDED(hRes) && 1 == cProp && lpCustomFlag && PT_LONG == PROP_TYPE(lpCustomFlag->ulPropTag))
				{
					LPSPropProblemArray lpProbArray2 = nullptr;
					// Clear the INSP_ONEOFFFLAGS bits so OL doesn't look for the props we deleted
					lpCustomFlag->Value.l = lpCustomFlag->Value.l & ~INSP_ONEOFFFLAGS;
					if (bRemovePropDef)
					{
						lpCustomFlag->Value.l = lpCustomFlag->Value.l & ~INSP_PROPDEFINITION;
					}

					hRes = EC_MAPI(lpMessage->SetProps(1, lpCustomFlag, &lpProbArray2));
					if (hRes == S_OK && lpProbArray2)
					{
						output::DebugPrint(
							output::dbgLevel::NamedProp,
							L"RemoveOneOff - SetProps problem array:\n%ws\n",
							error::ProblemArrayToString(*lpProbArray2).c_str());
					}

					MAPIFreeBuffer(lpProbArray2);
				}

				hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
				if (SUCCEEDED(hRes))
				{
					output::DebugPrint(output::dbgLevel::NamedProp, L"RemoveOneOff - One-off properties removed.\n");
				}
			}

			MAPIFreeBuffer(lpProbArray);
		}

		MAPIFreeBuffer(lpTags);
		return hRes;
	}

	_Check_return_ HRESULT ResendMessages(_In_ LPMAPIFOLDER lpFolder, _In_ HWND hWnd)
	{
		LPMAPITABLE lpContentsTable = nullptr;
		LPSRowSet pRows = nullptr;

		// You define a SPropTagArray array here using the SizedSPropTagArray Macro
		// This enum will allows you to access portions of the array by a name instead of a number.
		// If more tags are added to the array, appropriate constants need to be added to the enum.
		enum
		{
			ePR_ENTRYID,
			NUM_COLS
		};
		// These tags represent the message information we would like to pick up
		static const SizedSPropTagArray(NUM_COLS, sptCols) = {NUM_COLS, {PR_ENTRYID}};

		if (!lpFolder) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(lpFolder->GetContentsTable(0, &lpContentsTable));

		if (lpContentsTable)
		{
			hRes = EC_MAPI(HrQueryAllRows(
				lpContentsTable,
				LPSPropTagArray(&sptCols),
				nullptr, // restriction...we're not using this parameter
				nullptr, // sort order...we're not using this parameter
				0,
				&pRows));

			if (pRows)
			{
				if (SUCCEEDED(hRes))
				{
					for (ULONG i = 0; i < pRows->cRows; i++)
					{
						const auto& bin = mapi::getBin(pRows->aRow[i].lpProps[ePR_ENTRYID]);
						auto lpMessage = CallOpenEntry<LPMESSAGE>(
							nullptr,
							nullptr,
							lpFolder,
							nullptr,
							bin.cb,
							reinterpret_cast<LPENTRYID>(bin.lpb),
							nullptr,
							MAPI_BEST_ACCESS,
							nullptr);
						if (lpMessage)
						{
							hRes = EC_H(ResendSingleMessage(lpFolder, lpMessage, hWnd));
							lpMessage->Release();
						}
					}
				}
			}
		}

		if (pRows) FreeProws(pRows);
		if (lpContentsTable) lpContentsTable->Release();
		return hRes;
	}

	_Check_return_ HRESULT ResendSingleMessage(_In_ LPMAPIFOLDER lpFolder, _In_ LPSBinary MessageEID, _In_ HWND hWnd)
	{
		if (!lpFolder || !MessageEID) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		auto lpMessage = CallOpenEntry<LPMESSAGE>(
			nullptr,
			nullptr,
			lpFolder,
			nullptr,
			MessageEID->cb,
			reinterpret_cast<LPENTRYID>(MessageEID->lpb),
			nullptr,
			MAPI_BEST_ACCESS,
			nullptr);
		if (lpMessage)
		{
			hRes = EC_H(ResendSingleMessage(lpFolder, lpMessage, hWnd));
			lpMessage->Release();
		}

		return hRes;
	}

	_Check_return_ HRESULT ResendSingleMessage(_In_ LPMAPIFOLDER lpFolder, _In_ LPMESSAGE lpMessage, _In_ HWND hWnd)
	{
		auto hResRet = S_OK;
		LPATTACH lpAttach = nullptr;
		LPMESSAGE lpAttachMsg = nullptr;
		LPMAPITABLE lpAttachTable = nullptr;
		LPSRowSet pRows = nullptr;
		LPMESSAGE lpNewMessage = nullptr;
		LPSPropTagArray lpsMessageTags = nullptr;
		LPSPropProblemArray lpsPropProbs = nullptr;
		SPropValue sProp = {};

		enum
		{
			atPR_ATTACH_METHOD,
			atPR_ATTACH_NUM,
			atPR_DISPLAY_NAME,
			atNUM_COLS
		};

		static const SizedSPropTagArray(atNUM_COLS, atCols) = //
			{atNUM_COLS,
			 {
				 PR_ATTACH_METHOD, //
				 PR_ATTACH_NUM, //
				 PR_DISPLAY_NAME_W //
			 }};

		static const SizedSPropTagArray(2, atObjs) = //
			{2,
			 {
				 PR_MESSAGE_RECIPIENTS, //
				 PR_MESSAGE_ATTACHMENTS //
			 }};

		if (!lpMessage || !lpFolder) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(output::dbgLevel::Generic, L"ResendSingleMessage: Checking message for embedded messages\n");

		auto hRes = EC_MAPI(lpMessage->GetAttachmentTable(NULL, &lpAttachTable));

		if (lpAttachTable)
		{
			hRes = EC_MAPI(lpAttachTable->SetColumns(LPSPropTagArray(&atCols), TBL_BATCH));

			// Now we iterate through each of the attachments
			if (!FAILED(hRes))
			{
				for (;;)
				{
					// Remember the first error code we hit so it will bubble up
					if (FAILED(hRes) && SUCCEEDED(hResRet)) hResRet = hRes;
					if (pRows) FreeProws(pRows);
					pRows = nullptr;
					hRes = EC_MAPI(lpAttachTable->QueryRows(1, NULL, &pRows));
					if (FAILED(hRes)) break;
					if (!pRows || !pRows->cRows) break;

					if (ATTACH_EMBEDDED_MSG == pRows->aRow->lpProps[atPR_ATTACH_METHOD].Value.l)
					{
						output::DebugPrint(output::dbgLevel::Generic, L"Found an embedded message to resend.\n");

						if (lpAttach) lpAttach->Release();
						lpAttach = nullptr;
						hRes = EC_MAPI(lpMessage->OpenAttach(
							pRows->aRow->lpProps[atPR_ATTACH_NUM].Value.l,
							nullptr,
							MAPI_BEST_ACCESS,
							static_cast<LPATTACH*>(&lpAttach)));
						if (!lpAttach) continue;

						if (lpAttachMsg) lpAttachMsg->Release();
						lpAttachMsg = nullptr;
						hRes = EC_MAPI(lpAttach->OpenProperty(
							PR_ATTACH_DATA_OBJ,
							const_cast<LPIID>(&IID_IMessage),
							0,
							MAPI_MODIFY,
							reinterpret_cast<LPUNKNOWN*>(&lpAttachMsg)));
						if (hRes == MAPI_E_INTERFACE_NOT_SUPPORTED)
						{
							CHECKHRESMSG(hRes, IDS_ATTNOTEMBEDDEDMSG);
							continue;
						}

						if (FAILED(hRes)) continue;

						output::DebugPrint(output::dbgLevel::Generic, L"Message opened.\n");

						if (strings::CheckStringProp(&pRows->aRow->lpProps[atPR_DISPLAY_NAME], PT_UNICODE))
						{
							output::DebugPrint(
								output::dbgLevel::Generic,
								L"Resending \"%ws\"\n",
								pRows->aRow->lpProps[atPR_DISPLAY_NAME].Value.lpszW);
						}

						output::DebugPrint(output::dbgLevel::Generic, L"Creating new message.\n");
						if (lpNewMessage) lpNewMessage->Release();
						lpNewMessage = nullptr;
						hRes = EC_MAPI(lpFolder->CreateMessage(nullptr, 0, &lpNewMessage));
						if (FAILED(hRes) || !lpNewMessage) continue;

						hRes = EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));
						if (FAILED(hRes)) continue;

						// Copy all the transmission properties
						output::DebugPrint(output::dbgLevel::Generic, L"Getting list of properties.\n");
						MAPIFreeBuffer(lpsMessageTags);
						lpsMessageTags = nullptr;
						hRes = EC_MAPI(lpAttachMsg->GetPropList(0, &lpsMessageTags));
						if (FAILED(hRes) || !lpsMessageTags) continue;

						output::DebugPrint(output::dbgLevel::Generic, L"Copying properties to new message.\n");
						if (SUCCEEDED(hRes))
						{
							for (ULONG ulProp = 0; ulProp < lpsMessageTags->cValues; ulProp++)
							{
								// it would probably be quicker to use this loop to construct an array of properties
								// we desire to copy, and then pass that array to GetProps and then SetProps
								if (FIsTransmittable(getTag(lpsMessageTags, ulProp)))
								{
									LPSPropValue lpProp = nullptr;
									output::DebugPrint(
										output::dbgLevel::Generic, L"Copying 0x%08X\n", getTag(lpsMessageTags, ulProp));
									hRes = WC_MAPI(HrGetOnePropEx(
										lpAttachMsg, getTag(lpsMessageTags, ulProp), fMapiUnicode, &lpProp));

									if (SUCCEEDED(hRes))
									{
										hRes = WC_MAPI(HrSetOneProp(lpNewMessage, lpProp));
									}

									MAPIFreeBuffer(lpProp);
								}
							}
						}

						if (FAILED(hRes)) continue;

						hRes = EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));
						if (FAILED(hRes)) continue;

						output::DebugPrint(
							output::dbgLevel::Generic, L"Copying recipients and attachments to new message.\n");

						auto lpProgress = mapiui::GetMAPIProgress(L"IMAPIProp::CopyProps", hWnd); // STRING_OK

						hRes = EC_MAPI(lpAttachMsg->CopyProps(
							LPSPropTagArray(&atObjs),
							lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
							lpProgress,
							&IID_IMessage,
							lpNewMessage,
							lpProgress ? MAPI_DIALOG : 0,
							&lpsPropProbs));

						if (lpProgress) lpProgress->Release();

						if (lpsPropProbs)
						{
							EC_PROBLEMARRAY(lpsPropProbs);
							MAPIFreeBuffer(lpsPropProbs);
							lpsPropProbs = nullptr;
							continue;
						}

						sProp.dwAlignPad = 0;
						sProp.ulPropTag = PR_DELETE_AFTER_SUBMIT;
						sProp.Value.b = true;

						output::DebugPrint(output::dbgLevel::Generic, L"Setting PR_DELETE_AFTER_SUBMIT to true.\n");
						hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));
						if (FAILED(hRes)) continue;

						SPropTagArray sPropTagArray = {1, PR_SENTMAIL_ENTRYID};

						output::DebugPrint(output::dbgLevel::Generic, L"Deleting PR_SENTMAIL_ENTRYID\n");
						hRes = EC_MAPI(lpNewMessage->DeleteProps(&sPropTagArray, nullptr));
						if (FAILED(hRes)) continue;

						hRes = EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));
						if (FAILED(hRes)) continue;

						output::DebugPrint(output::dbgLevel::Generic, L"Submitting new message.\n");
						hRes = EC_MAPI(lpNewMessage->SubmitMessage(0));
						if (FAILED(hRes)) continue;
					}
					else
					{
						output::DebugPrint(output::dbgLevel::Generic, L"Attachment is not an embedded message.\n");
					}
				}
			}
		}

		MAPIFreeBuffer(lpsMessageTags);
		if (lpNewMessage) lpNewMessage->Release();
		if (lpAttachMsg) lpAttachMsg->Release();
		if (lpAttach) lpAttach->Release();
		if (pRows) FreeProws(pRows);
		if (lpAttachTable) lpAttachTable->Release();
		if (FAILED(hResRet)) return hResRet;
		return hRes;
	}

	_Check_return_ HRESULT ResetPermissionsOnItems(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpMAPIFolder)
	{
		LPSRowSet pRows = nullptr;
		auto hRes = S_OK;
		auto hResOverall = S_OK;
		ULONG iItemCount = 0;
		LPMAPITABLE lpContentsTable = nullptr;
		LPMESSAGE lpMessage = nullptr;

		enum
		{
			eidPR_ENTRYID,
			eidNUM_COLS
		};

		static const SizedSPropTagArray(eidNUM_COLS, eidCols) = {
			eidNUM_COLS,
			{PR_ENTRYID},
		};

		if (!lpMDB || !lpMAPIFolder) return MAPI_E_INVALID_PARAMETER;

		// We pass through this code twice, once for regular contents, once for associated contents
		for (auto i = 0; i <= 1; i++)
		{
			const auto ulFlags = (i == 1 ? MAPI_ASSOCIATED : NULL) | fMapiUnicode;

			if (lpContentsTable) lpContentsTable->Release();
			lpContentsTable = nullptr;
			// Get the table of contents of the folder
			hRes = EC_MAPI(lpMAPIFolder->GetContentsTable(ulFlags, &lpContentsTable));

			if (SUCCEEDED(hRes) && lpContentsTable)
			{
				hRes = EC_MAPI(lpContentsTable->SetColumns(LPSPropTagArray(&eidCols), TBL_BATCH));

				if (SUCCEEDED(hRes))
				{
					// go to the first row
					hRes = EC_MAPI(lpContentsTable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));
				}

				// get rows and delete PR_NT_SECURITY_DESCRIPTOR
				if (!FAILED(hRes))
				{
					for (;;)
					{
						if (pRows) FreeProws(pRows);
						pRows = nullptr;
						// Pull back a sizable block of rows to modify
						hRes = EC_MAPI(lpContentsTable->QueryRows(200, NULL, &pRows));
						if (FAILED(hRes) || !pRows || !pRows->cRows) break;

						for (ULONG iCurPropRow = 0; iCurPropRow < pRows->cRows; iCurPropRow++)
						{
							if (lpMessage) lpMessage->Release();
							const auto& bin = mapi::getBin(pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID]);
							lpMessage = CallOpenEntry<LPMESSAGE>(
								lpMDB,
								nullptr,
								nullptr,
								nullptr,
								bin.cb,
								reinterpret_cast<LPENTRYID>(bin.lpb),
								nullptr,
								MAPI_BEST_ACCESS,
								nullptr);
							if (lpMessage)
							{

								hRes = WC_H(DeleteProperty(lpMessage, PR_NT_SECURITY_DESCRIPTOR));
								if (FAILED(hRes))
								{
									hResOverall = hRes;
									continue;
								}
							}
							iItemCount++;
						}
					}
				}

				output::DebugPrint(
					output::dbgLevel::Generic, L"ResetPermissionsOnItems reset permissions on %u items\n", iItemCount);
			}
		}

		if (pRows) FreeProws(pRows);
		if (lpMessage) lpMessage->Release();
		if (lpContentsTable) lpContentsTable->Release();
		if (S_OK != hResOverall) return hResOverall;
		return hRes;
	}

	// This function creates a new message based in lpFolder
	// Then sends the message
	_Check_return_ HRESULT SendTestMessage(
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMAPIFOLDER lpFolder,
		_In_ const std::wstring& szRecipient,
		_In_ const std::wstring& szBody,
		_In_ const std::wstring& szSubject,
		_In_ const std::wstring& szClass)
	{
		LPMESSAGE lpNewMessage = nullptr;

		if (!lpMAPISession || !lpFolder) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(lpFolder->CreateMessage(
			nullptr, // default interface
			0, // flags
			&lpNewMessage));

		if (lpNewMessage)
		{
			auto sProp = SPropValue{PR_DELETE_AFTER_SUBMIT};
			sProp.Value.b = true;

			output::DebugPrint(output::dbgLevel::Generic, L"Setting PR_DELETE_AFTER_SUBMIT to true.\n");
			hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));

			if (SUCCEEDED(hRes))
			{
				sProp.dwAlignPad = 0;
				sProp.ulPropTag = PR_BODY_W;
				sProp.Value.lpszW = const_cast<LPWSTR>(szBody.c_str());

				output::DebugPrint(output::dbgLevel::Generic, L"Setting PR_BODY to %ws.\n", szBody.c_str());
				hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));
			}

			if (SUCCEEDED(hRes))
			{
				sProp.dwAlignPad = 0;
				sProp.ulPropTag = PR_SUBJECT_W;
				sProp.Value.lpszW = const_cast<LPWSTR>(szSubject.c_str());

				output::DebugPrint(output::dbgLevel::Generic, L"Setting PR_SUBJECT to %ws.\n", szSubject.c_str());
				hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));
			}

			if (SUCCEEDED(hRes))
			{
				sProp.dwAlignPad = 0;
				sProp.ulPropTag = PR_MESSAGE_CLASS_W;
				sProp.Value.lpszW = const_cast<LPWSTR>(szClass.c_str());

				output::DebugPrint(output::dbgLevel::Generic, L"Setting PR_MESSAGE_CLASS to %ws.\n", szSubject.c_str());
				hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));
			}

			if (SUCCEEDED(hRes))
			{
				SPropTagArray sPropTagArray = {1, PR_SENTMAIL_ENTRYID};

				output::DebugPrint(output::dbgLevel::Generic, L"Deleting PR_SENTMAIL_ENTRYID\n");
				hRes = EC_MAPI(lpNewMessage->DeleteProps(&sPropTagArray, nullptr));
			}

			if (SUCCEEDED(hRes))
			{
				output::DebugPrint(output::dbgLevel::Generic, L"Adding recipient: %ws.\n", szRecipient.c_str());
				hRes = EC_H(ab::AddRecipient(lpMAPISession, lpNewMessage, szRecipient, MAPI_TO));
			}

			if (SUCCEEDED(hRes))
			{
				output::DebugPrint(output::dbgLevel::Generic, L"Submitting message\n");
				hRes = EC_MAPI(lpNewMessage->SubmitMessage(NULL));
			}
		}

		if (lpNewMessage) lpNewMessage->Release();
		return hRes;
	}

	_Check_return_ HRESULT WrapStreamForRTF(
		_In_ LPSTREAM lpCompressedRTFStream,
		bool bUseWrapEx,
		ULONG ulFlags,
		ULONG ulInCodePage,
		ULONG ulOutCodePage,
		_Deref_out_ LPSTREAM* lpUncompressedRTFStream,
		_Out_opt_ ULONG* pulStreamFlags)
	{
		if (pulStreamFlags) *pulStreamFlags = {};
		if (!lpCompressedRTFStream || !lpUncompressedRTFStream) return MAPI_E_INVALID_PARAMETER;
		auto hRes = S_OK;

		if (!bUseWrapEx)
		{
			hRes = WC_MAPI(WrapCompressedRTFStream(lpCompressedRTFStream, ulFlags, lpUncompressedRTFStream));
		}
		else
		{
			RTF_WCSINFO wcsinfo = {};
			RTF_WCSRETINFO retinfo = {};

			retinfo.size = sizeof(RTF_WCSRETINFO);

			wcsinfo.size = sizeof(RTF_WCSINFO);
			wcsinfo.ulFlags = ulFlags;
			wcsinfo.ulInCodePage = ulInCodePage; // Get ulCodePage from PR_INTERNET_CPID on the IMessage
			wcsinfo.ulOutCodePage = ulOutCodePage; // Desired code page for return

			hRes =
				WC_MAPI(WrapCompressedRTFStreamEx(lpCompressedRTFStream, &wcsinfo, lpUncompressedRTFStream, &retinfo));
			if (pulStreamFlags) *pulStreamFlags = retinfo.ulStreamFlags;
		}

		return hRes;
	}

	_Check_return_ HRESULT CopyNamedProps(
		_In_ LPMAPIPROP lpSource,
		_In_ LPGUID lpPropSetGUID,
		bool bDoMove,
		bool bDoNoReplace,
		_In_ LPMAPIPROP lpTarget,
		_In_ HWND hWnd)
	{
		if (!lpSource || !lpTarget) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		auto lpPropTags = GetNamedPropsByGUID(lpSource, lpPropSetGUID);
		if (lpPropTags)
		{
			LPSPropProblemArray lpProblems = nullptr;
			ULONG ulFlags = 0;
			if (bDoMove) ulFlags |= MAPI_MOVE;
			if (bDoNoReplace) ulFlags |= MAPI_NOREPLACE;

			auto lpProgress = mapiui::GetMAPIProgress(L"IMAPIProp::CopyProps", hWnd); // STRING_OK

			if (lpProgress) ulFlags |= MAPI_DIALOG;

			hRes = EC_MAPI(lpSource->CopyProps(
				lpPropTags,
				lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
				lpProgress,
				&IID_IMAPIProp,
				lpTarget,
				ulFlags,
				&lpProblems));

			if (lpProgress) lpProgress->Release();

			EC_PROBLEMARRAY(lpProblems);
			MAPIFreeBuffer(lpProblems);
		}

		MAPIFreeBuffer(lpPropTags);

		return hRes;
	}

	_Check_return_ LPSPropTagArray GetNamedPropsByGUID(_In_ LPMAPIPROP lpSource, _In_ LPGUID lpPropSetGUID)
	{
		if (!lpSource || !lpPropSetGUID) return nullptr;

		LPSPropTagArray lpAllProps = nullptr;
		LPSPropTagArray lpFilteredProps = nullptr;

		const auto hRes = WC_MAPI(lpSource->GetPropList(0, &lpAllProps));
		if (hRes == S_OK && lpAllProps)
		{
			const auto names = cache::GetNamesFromIDs(lpSource, lpAllProps, 0);
			if (!names.empty())
			{
				ULONG ulNumProps = 0; // count of props that match our guid
				for (const auto& name : names)
				{
					if (cache::namedPropCacheEntry::valid(name) && name->getPropID() > 0x7FFF &&
						::IsEqualGUID(*name->getMapiNameId()->lpguid, *lpPropSetGUID))
					{
						ulNumProps++;
					}
				}

				lpFilteredProps = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(ulNumProps));
				if (lpFilteredProps)
				{
					lpFilteredProps->cValues = 0;

					for (const auto& name : names)
					{
						if (cache::namedPropCacheEntry::valid(name) && name->getPropID() > 0x7FFF &&
							::IsEqualGUID(*name->getMapiNameId()->lpguid, *lpPropSetGUID))
						{
							setTag(lpFilteredProps, lpFilteredProps->cValues) = name->getPropID();
							lpFilteredProps->cValues++;
						}
					}
				}
			}
		}

		MAPIFreeBuffer(lpAllProps);
		return lpFilteredProps;
	}

	_Check_return_ DWORD ComputeStoreHash(
		ULONG cbStoreEID,
		_In_count_(cbStoreEID) LPBYTE pbStoreEID,
		_In_opt_z_ LPCSTR pszFileName,
		_In_opt_z_ LPCWSTR pwzFileName,
		bool bPublicStore)
	{
		DWORD dwHash = 0;

		if (!cbStoreEID || !pbStoreEID) return dwHash;
		// We shouldn't see both of these at the same time.
		if (pszFileName && pwzFileName) return dwHash;

		// Get the Store Entry ID
		// pbStoreEID is a pointer to the Entry ID
		// cbStoreEID is the size in bytes of the Entry ID
		auto pdw = reinterpret_cast<LPDWORD>(pbStoreEID);
		const auto cdw = cbStoreEID / sizeof(DWORD);

		for (ULONG i = 0; i < cdw; i++)
		{
			dwHash = (dwHash << 5) + dwHash + *pdw++;
		}

		auto pb = reinterpret_cast<LPBYTE>(pdw);
		const auto cb = cbStoreEID % sizeof(DWORD);

		for (ULONG i = 0; i < cb; i++)
		{
			dwHash = (dwHash << 5) + dwHash + *pb++;
		}

		if (bPublicStore)
		{
			output::DebugPrint(
				output::dbgLevel::Generic, L"ComputeStoreHash, hash (before adding .PUB) = 0x%08X\n", dwHash);
			// augment to make sure it is unique else could be same as the private store
			dwHash = (dwHash << 5) + dwHash + 0x2E505542; // this is '.PUB'
		}

		if (pwzFileName || pszFileName)
			output::DebugPrint(
				output::dbgLevel::Generic, L"ComputeStoreHash, hash (before adding path) = 0x%08X\n", dwHash);

		// You may want to also include the store file name in the hash calculation
		// pszFileName and pwzFileName are NULL terminated strings with the path and filename of the store
		if (pwzFileName)
		{
			while (*pwzFileName)
			{
				dwHash = (dwHash << 5) + dwHash + *pwzFileName++;
			}
		}
		else if (pszFileName)
		{
			while (*pszFileName)
			{
				dwHash = (dwHash << 5) + dwHash + *pszFileName++;
			}
		}

		if (pwzFileName || pszFileName)
			output::DebugPrint(
				output::dbgLevel::Generic, L"ComputeStoreHash, hash (after adding path) = 0x%08X\n", dwHash);

		// dwHash now contains the hash to be used. It should be written in hex when building a URL.
		return dwHash;
	}

	HRESULT
	HrEmsmdbUIDFromStore(_In_ LPMAPISESSION pmsess, _In_ const MAPIUID* puidService, _Out_opt_ MAPIUID* pEmsmdbUID)
	{
		if (pEmsmdbUID) *pEmsmdbUID = {};
		if (!puidService) return MAPI_E_INVALID_PARAMETER;

		SRestriction mres = {};
		SPropValue mval = {};
		SRowSet* pRows = nullptr;
		LPSERVICEADMIN spSvcAdmin = nullptr;
		LPMAPITABLE spmtab = nullptr;

		enum
		{
			eEntryID = 0,
			eSectionUid,
			eMax
		};
		static const SizedSPropTagArray(eMax, tagaCols) = {
			eMax,
			{
				PR_ENTRYID,
				PR_EMSMDB_SECTION_UID,
			}};

		auto hRes = EC_MAPI(pmsess->AdminServices(0, static_cast<LPSERVICEADMIN*>(&spSvcAdmin)));
		if (spSvcAdmin)
		{
			hRes = EC_MAPI(spSvcAdmin->GetMsgServiceTable(0, &spmtab));
			if (spmtab)
			{
				hRes = EC_MAPI(spmtab->SetColumns(LPSPropTagArray(&tagaCols), TBL_BATCH));

				if (SUCCEEDED(hRes))
				{
					mres.rt = RES_PROPERTY;
					mres.res.resProperty.relop = RELOP_EQ;
					mres.res.resProperty.ulPropTag = PR_SERVICE_UID;
					mres.res.resProperty.lpProp = &mval;
					mval.ulPropTag = PR_SERVICE_UID;
					mapi::setBin(mval) = {
						sizeof *puidService, reinterpret_cast<LPBYTE>(const_cast<MAPIUID*>(puidService))};

					hRes = EC_MAPI(spmtab->Restrict(&mres, 0));
				}

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI(spmtab->QueryRows(10, 0, &pRows));
				}

				if (SUCCEEDED(hRes) && pRows && pRows->cRows)
				{
					const auto pRow = &pRows->aRow[0];

					if (pEmsmdbUID && pRow)
					{
						const auto& bin = mapi::getBin(pRow->lpProps[eSectionUid]);
						if (PR_EMSMDB_SECTION_UID == pRow->lpProps[eSectionUid].ulPropTag &&
							bin.cb == sizeof *pEmsmdbUID)
						{
							memcpy(pEmsmdbUID, bin.lpb, sizeof *pEmsmdbUID);
						}
					}
				}

				FreeProws(pRows);
			}

			if (spmtab) spmtab->Release();
		}

		if (spSvcAdmin) spSvcAdmin->Release();
		return hRes;
	}

	bool FExchangePrivateStore(_In_ LPMAPIUID lpmapiuid) noexcept
	{
		if (!lpmapiuid) return false;
		return IsEqualMAPIUID(lpmapiuid, LPMAPIUID(pbExchangeProviderPrimaryUserGuid));
	}

	bool FExchangePublicStore(_In_ LPMAPIUID lpmapiuid) noexcept
	{
		if (!lpmapiuid) return false;
		return IsEqualMAPIUID(lpmapiuid, LPMAPIUID(pbExchangeProviderPublicGuid));
	}

	LPSBinary GetEntryIDFromMDB(LPMDB lpMDB, ULONG ulPropTag)
	{
		if (!lpMDB) return {};
		LPSPropValue lpEIDProp = nullptr;

		const auto hRes = WC_MAPI(HrGetOneProp(lpMDB, ulPropTag, &lpEIDProp));

		auto eid = LPSBinary{};
		if (SUCCEEDED(hRes) && lpEIDProp)
		{
			eid = CopySBinary(&mapi::getBin(lpEIDProp));
		}

		MAPIFreeBuffer(lpEIDProp);
		return eid;
	}

	LPSBinary GetMVEntryIDFromInboxByIndex(LPMDB lpMDB, ULONG ulPropTag, ULONG ulIndex)
	{
		if (!lpMDB) return {};

		auto eid = LPSBinary{};
		auto lpInbox = GetInbox(lpMDB);
		if (lpInbox)
		{
			LPSPropValue lpEIDProp = nullptr;
			WC_MAPI_S(HrGetOneProp(lpInbox, ulPropTag, &lpEIDProp));
			if (lpEIDProp && PT_MV_BINARY == PROP_TYPE(lpEIDProp->ulPropTag) &&
				ulIndex < lpEIDProp->Value.MVbin.cValues && lpEIDProp->Value.MVbin.lpbin[ulIndex].cb > 0)
			{
				eid = CopySBinary(&lpEIDProp->Value.MVbin.lpbin[ulIndex]);
			}

			MAPIFreeBuffer(lpEIDProp);
			lpInbox->Release();
		}

		return eid;
	}

	LPSBinary GetDefaultFolderEID(_In_ ULONG ulFolder, _In_ LPMDB lpMDB)
	{
		if (!lpMDB) return {};

		switch (ulFolder)
		{
		case DEFAULT_CALENDAR:
			return GetSpecialFolderEID(lpMDB, PR_IPM_APPOINTMENT_ENTRYID);
		case DEFAULT_CONTACTS:
			return GetSpecialFolderEID(lpMDB, PR_IPM_CONTACT_ENTRYID);
		case DEFAULT_JOURNAL:
			return GetSpecialFolderEID(lpMDB, PR_IPM_JOURNAL_ENTRYID);
		case DEFAULT_NOTES:
			return GetSpecialFolderEID(lpMDB, PR_IPM_NOTE_ENTRYID);
		case DEFAULT_TASKS:
			return GetSpecialFolderEID(lpMDB, PR_IPM_TASK_ENTRYID);
		case DEFAULT_REMINDERS:
			return GetSpecialFolderEID(lpMDB, PR_REM_ONLINE_ENTRYID);
		case DEFAULT_DRAFTS:
			return GetSpecialFolderEID(lpMDB, PR_IPM_DRAFTS_ENTRYID);
		case DEFAULT_SENTITEMS:
			return GetEntryIDFromMDB(lpMDB, PR_IPM_SENTMAIL_ENTRYID);
		case DEFAULT_OUTBOX:
			return GetEntryIDFromMDB(lpMDB, PR_IPM_OUTBOX_ENTRYID);
		case DEFAULT_DELETEDITEMS:
			return GetEntryIDFromMDB(lpMDB, PR_IPM_WASTEBASKET_ENTRYID);
		case DEFAULT_FINDER:
			return GetEntryIDFromMDB(lpMDB, PR_FINDER_ENTRYID);
		case DEFAULT_IPM_SUBTREE:
			return GetEntryIDFromMDB(lpMDB, PR_IPM_SUBTREE_ENTRYID);
		case DEFAULT_INBOX:
			return GetInboxEntryId(lpMDB);
		case DEFAULT_LOCALFREEBUSY:
			return GetMVEntryIDFromInboxByIndex(lpMDB, PR_FREEBUSY_ENTRYIDS, 3);
		case DEFAULT_CONFLICTS:
			return GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 0);
		case DEFAULT_SYNCISSUES:
			return GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 1);
		case DEFAULT_LOCALFAILURES:
			return GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 2);
		case DEFAULT_SERVERFAILURES:
			return GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 3);
		case DEFAULT_JUNKMAIL:
			return GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 4);
		}

		return {};
	}

	LPMAPIFOLDER OpenDefaultFolder(_In_ ULONG ulFolder, _In_ LPMDB lpMDB)
	{
		if (!lpMDB) return nullptr;

		LPMAPIFOLDER lpFolder = nullptr;

		const auto eid = GetDefaultFolderEID(ulFolder, lpMDB);
		if (eid)
		{
			lpFolder =
				CallOpenEntry<LPMAPIFOLDER>(lpMDB, nullptr, nullptr, nullptr, eid, nullptr, MAPI_BEST_ACCESS, nullptr);
			MAPIFreeBuffer(eid);
		}

		return lpFolder;
	}

	ULONG g_DisplayNameProps[] = {
		PR_DISPLAY_NAME_W,
		CHANGE_PROP_TYPE(PR_MAILBOX_OWNER_NAME, PT_UNICODE),
		PR_SUBJECT_W,
	};

	std::wstring GetTitle(LPMAPIPROP lpMAPIProp)
	{
		std::wstring szTitle;
		LPSPropValue lpProp = nullptr;
		auto bFoundName = false;

		if (!lpMAPIProp) return szTitle;

		// Get a property for the title bar
		for (ULONG i = 0; !bFoundName && i < _countof(g_DisplayNameProps); i++)
		{
			WC_MAPI_S(HrGetOneProp(lpMAPIProp, g_DisplayNameProps[i], &lpProp));

			if (lpProp)
			{
				if (strings::CheckStringProp(lpProp, PT_UNICODE))
				{
					szTitle = lpProp->Value.lpszW;
					bFoundName = true;
				}

				MAPIFreeBuffer(lpProp);
			}
		}

		if (!bFoundName)
		{
			szTitle = strings::loadstring(IDS_DISPLAYNAMENOTFOUND);
		}

		return szTitle;
	}

	bool UnwrapContactEntryID(_In_ ULONG cbIn, _In_ LPBYTE lpbIn, _Out_ ULONG* lpcbOut, _Out_ LPBYTE* lppbOut) noexcept
	{
		if (lpcbOut) *lpcbOut = 0;
		if (lppbOut) *lppbOut = nullptr;

		if (cbIn < sizeof(DIR_ENTRYID)) return false;
		if (!lpcbOut || !lppbOut || !lpbIn) return false;

		const auto lpContabEID = reinterpret_cast<LPCONTAB_ENTRYID>(lpbIn);

		switch (lpContabEID->ulType)
		{
		case CONTAB_USER:
		case CONTAB_DISTLIST:
			if (cbIn >= sizeof(CONTAB_ENTRYID) && lpContabEID->cbeid)
			{
				*lpcbOut = lpContabEID->cbeid;
				*lppbOut = lpContabEID->abeid;
				return true;
			}
			break;
		case CONTAB_ROOT:
		case CONTAB_SUBROOT:
		case CONTAB_CONTAINER:
			*lpcbOut = cbIn - sizeof(DIR_ENTRYID);
			*lppbOut = lpbIn + sizeof(DIR_ENTRYID);
			return true;
		}

		return false;
	}

	// Augemented version of HrGetOneProp which allows passing flags to underlying GetProps
	// Useful for passing fMapiUnicode for unspecified string/stream types
	HRESULT
	HrGetOnePropEx(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ ULONG ulPropTag,
		_In_ ULONG ulFlags,
		_Out_ LPSPropValue* lppProp) noexcept
	{
		if (!lppProp) return MAPI_E_INVALID_PARAMETER;
		*lppProp = nullptr;
		ULONG cValues = 0;
		SPropTagArray tag = {1, {ulPropTag}};

		LPSPropValue lpProp = nullptr;
		auto hRes = lpMAPIProp->GetProps(&tag, ulFlags, &cValues, &lpProp);
		if (SUCCEEDED(hRes))
		{
			if (lpProp && PROP_TYPE(lpProp->ulPropTag) == PT_ERROR)
			{
				hRes = ResultFromScode(lpProp->Value.err);
				MAPIFreeBuffer(lpProp);
				lpProp = nullptr;
			}

			*lppProp = lpProp;
		}

		return hRes;
	}

	void ForceRop(_In_ LPMDB lpMDB)
	{
		LPSPropValue lpProp = nullptr;
		// Try to trigger a rop to get notifications going
		WC_MAPI_S(HrGetOneProp(lpMDB, PR_TEST_LINE_SPEED, &lpProp));
		// No need to worry about errors here - this is just to force rops
		MAPIFreeBuffer(lpProp);
	}

	_Check_return_ HRESULT
	HrDupPropset(int cprop, _In_count_(cprop) LPSPropValue rgprop, _In_ const VOID* parent, _In_ LPSPropValue* prgprop)
	{
		ULONG cb = NULL;

		// Find out how much memory we need
		auto hRes = EC_MAPI(ScCountProps(cprop, rgprop, &cb));

		if (SUCCEEDED(hRes) && cb)
		{
			// Obtain memory
			*prgprop = mapi::allocate<LPSPropValue>(cb, parent);
			if (*prgprop)
			{
				// Copy the properties
				hRes = EC_MAPI(ScCopyProps(cprop, rgprop, *prgprop, &cb));
			}
		}

		return hRes;
	}

	_Check_return_ STDAPI HrCopyRestriction(
		_In_ const _SRestriction* lpResSrc, // source restriction ptr
		_In_opt_ const void* lpObject, // ptr to existing MAPI buffer
		_In_ LPSRestriction* lppResDest) // dest restriction buffer ptr
	{
		if (!lppResDest) return MAPI_E_INVALID_PARAMETER;
		*lppResDest = nullptr;
		if (!lpResSrc) return S_OK;

		*lppResDest = mapi::allocate<LPSRestriction>(sizeof(SRestriction), lpObject);
		auto lpAllocationParent = lpObject ? lpObject : *lppResDest;
		const auto hRes =
			WC_H(HrCopyRestrictionArray(lpResSrc, const_cast<LPVOID>(lpAllocationParent), 1, *lppResDest));

		if (FAILED(hRes))
		{
			if (!lpObject) MAPIFreeBuffer(*lppResDest);
		}

		return hRes;
	}

	_Check_return_ HRESULT HrCopyRestrictionArray(
		_In_ const _SRestriction* lpResSrc, // source restriction
		_In_ const VOID* lpObject, // ptr to existing MAPI buffer
		ULONG cRes, // # elements in array
		_In_count_(cRes) LPSRestriction lpResDest) // destination restriction
	{
		if (!lpResSrc || !lpResDest || !lpObject) return MAPI_E_INVALID_PARAMETER;
		auto hRes = S_OK;

		for (ULONG i = 0; i < cRes; i++)
		{
			// Copy all the members over
			lpResDest[i] = lpResSrc[i];

			// Now fix up the pointers
			switch (lpResSrc[i].rt)
			{
				// Structures for these two types are identical
			case RES_AND:
			case RES_OR:
				if (lpResSrc[i].res.resAnd.cRes && lpResSrc[i].res.resAnd.lpRes)
				{
					if (lpResSrc[i].res.resAnd.cRes > ULONG_MAX / sizeof(SRestriction))
					{
						hRes = MAPI_E_CALL_FAILED;
						break;
					}

					lpResDest[i].res.resAnd.lpRes =
						mapi::allocate<LPSRestriction>(sizeof(SRestriction) * lpResSrc[i].res.resAnd.cRes, lpObject);
					auto lpAllocationParent = lpObject ? lpObject : lpResDest[i].res.resAnd.lpRes;

					hRes = WC_H(HrCopyRestrictionArray(
						lpResSrc[i].res.resAnd.lpRes,
						lpAllocationParent,
						lpResSrc[i].res.resAnd.cRes,
						lpResDest[i].res.resAnd.lpRes));
					if (FAILED(hRes)) break;
				}
				break;

				// Structures for these two types are identical
			case RES_NOT:
			case RES_COUNT:
				if (lpResSrc[i].res.resNot.lpRes)
				{
					lpResDest[i].res.resNot.lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), lpObject);

					hRes = WC_H(HrCopyRestrictionArray(
						lpResSrc[i].res.resNot.lpRes, lpObject, 1, lpResDest[i].res.resNot.lpRes));
					if (FAILED(hRes)) break;
				}
				break;

				// Structures for these two types are identical
			case RES_CONTENT:
			case RES_PROPERTY:
				if (lpResSrc[i].res.resContent.lpProp)
				{
					hRes = WC_MAPI(HrDupPropset(
						1,
						lpResSrc[i].res.resContent.lpProp,
						const_cast<LPVOID>(lpObject),
						&lpResDest[i].res.resContent.lpProp));
					if (FAILED(hRes)) break;
				}
				break;

			case RES_COMPAREPROPS:
			case RES_BITMASK:
			case RES_SIZE:
			case RES_EXIST:
				break; // Nothing to do.

			case RES_SUBRESTRICTION:
				if (lpResSrc[i].res.resSub.lpRes)
				{
					lpResDest[i].res.resSub.lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), lpObject);

					hRes = WC_H(HrCopyRestrictionArray(
						lpResSrc[i].res.resSub.lpRes, lpObject, 1, lpResDest[i].res.resSub.lpRes));
					if (FAILED(hRes)) break;
				}
				break;

				// Structures for these two types are identical
			case RES_COMMENT:
			case RES_ANNOTATION:
				if (lpResSrc[i].res.resComment.lpRes)
				{
					lpResDest[i].res.resComment.lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), lpObject);
					if (FAILED(hRes)) break;

					hRes = WC_H(HrCopyRestrictionArray(
						lpResSrc[i].res.resComment.lpRes, lpObject, 1, lpResDest[i].res.resComment.lpRes));
					if (FAILED(hRes)) break;
				}

				if (lpResSrc[i].res.resComment.cValues && lpResSrc[i].res.resComment.lpProp)
				{
					hRes = WC_MAPI(mapi::HrDupPropset(
						lpResSrc[i].res.resComment.cValues,
						lpResSrc[i].res.resComment.lpProp,
						const_cast<LPVOID>(lpObject),
						&lpResDest[i].res.resComment.lpProp));
					if (FAILED(hRes)) break;
				}
				break;

			default:
				hRes = MAPI_E_INVALID_PARAMETER;
				break;
			}
		}

		return hRes;
	}

	// swiped from EDK rules sample
	_Check_return_ STDAPI HrCopyActions(
		_In_ LPACTIONS lpActsSrc, // source action ptr
		_In_ const VOID* parent, // ptr to existing MAPI buffer
		_In_ LPACTIONS* lppActsDst) // ptr to destination ACTIONS buffer
	{
		if (!lpActsSrc || !lppActsDst) return MAPI_E_INVALID_PARAMETER;
		if (lpActsSrc->cActions <= 0 || lpActsSrc->lpAction == nullptr) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;

		*lppActsDst = mapi::allocate<LPACTIONS>(sizeof(ACTIONS), parent);
		const auto lpAllocationParent = parent ? parent : *lppActsDst;
		// no short circuit returns after here

		const auto lpActsDst = *lppActsDst;
		*lpActsDst = *lpActsSrc;
		lpActsDst->lpAction = nullptr;

		lpActsDst->lpAction = mapi::allocate<LPACTION>(sizeof(ACTION) * lpActsDst->cActions, lpAllocationParent);
		if (lpActsDst->lpAction)
		{
			// Initialize acttype values for all members of the array to a value
			// that will not cause deallocation errors should the copy fail.
			for (ULONG i = 0; i < lpActsDst->cActions; i++)
				lpActsDst->lpAction[i].acttype = OP_BOUNCE;

			// Now actually copy all the members of the array.
			for (ULONG i = 0; i < lpActsDst->cActions; i++)
			{
				auto lpActDst = &lpActsDst->lpAction[i];
				auto lpActSrc = &lpActsSrc->lpAction[i];

				*lpActDst = *lpActSrc;

				switch (lpActSrc->acttype)
				{
				case OP_MOVE: // actMoveCopy
				case OP_COPY:
					if (lpActDst->actMoveCopy.cbStoreEntryId && lpActDst->actMoveCopy.lpStoreEntryId)
					{
						lpActDst->actMoveCopy.lpStoreEntryId =
							mapi::allocate<LPENTRYID>(lpActDst->actMoveCopy.cbStoreEntryId, lpAllocationParent);

						memcpy(
							lpActDst->actMoveCopy.lpStoreEntryId,
							lpActSrc->actMoveCopy.lpStoreEntryId,
							lpActSrc->actMoveCopy.cbStoreEntryId);
					}

					if (lpActDst->actMoveCopy.cbFldEntryId && lpActDst->actMoveCopy.lpFldEntryId)
					{
						lpActDst->actMoveCopy.lpFldEntryId =
							mapi::allocate<LPENTRYID>(lpActDst->actMoveCopy.cbFldEntryId, lpAllocationParent);

						memcpy(
							lpActDst->actMoveCopy.lpFldEntryId,
							lpActSrc->actMoveCopy.lpFldEntryId,
							lpActSrc->actMoveCopy.cbFldEntryId);
					}
					break;

				case OP_REPLY: // actReply
				case OP_OOF_REPLY:
					if (lpActDst->actReply.cbEntryId && lpActDst->actReply.lpEntryId)
					{
						lpActDst->actReply.lpEntryId =
							mapi::allocate<LPENTRYID>(lpActDst->actReply.cbEntryId, lpAllocationParent);

						memcpy(
							lpActDst->actReply.lpEntryId, lpActSrc->actReply.lpEntryId, lpActSrc->actReply.cbEntryId);
					}
					break;

				case OP_DEFER_ACTION: // actDeferAction
					if (lpActSrc->actDeferAction.pbData && lpActSrc->actDeferAction.cbData)
					{
						lpActDst->actDeferAction.pbData =
							mapi::allocate<LPBYTE>(lpActDst->actDeferAction.cbData, lpAllocationParent);

						memcpy(
							lpActDst->actDeferAction.pbData,
							lpActSrc->actDeferAction.pbData,
							lpActDst->actDeferAction.cbData);
					}
					break;

				case OP_FORWARD: // lpadrlist
				case OP_DELEGATE:
					lpActDst->lpadrlist = nullptr;

					if (lpActSrc->lpadrlist && lpActSrc->lpadrlist->cEntries)
					{
						lpActDst->lpadrlist =
							mapi::allocate<LPADRLIST>(CbADRLIST(lpActSrc->lpadrlist), lpAllocationParent);
						lpActDst->lpadrlist->cEntries = lpActSrc->lpadrlist->cEntries;

						// Initialize the new ADRENTRYs and validate cValues.
						for (ULONG j = 0; j < lpActSrc->lpadrlist->cEntries; j++)
						{
							lpActDst->lpadrlist->aEntries[j] = lpActSrc->lpadrlist->aEntries[j];
							lpActDst->lpadrlist->aEntries[j].rgPropVals = nullptr;

							if (lpActDst->lpadrlist->aEntries[j].cValues == 0)
							{
								hRes = MAPI_E_INVALID_PARAMETER;
								break;
							}
						}

						// Copy the rgPropVals.
						for (ULONG j = 0; j < lpActSrc->lpadrlist->cEntries; j++)
						{
							hRes = WC_MAPI(HrDupPropset(
								lpActDst->lpadrlist->aEntries[j].cValues,
								lpActSrc->lpadrlist->aEntries[j].rgPropVals,
								parent,
								&lpActDst->lpadrlist->aEntries[j].rgPropVals));
							if (FAILED(hRes)) break;
						}
					}
					break;

				case OP_TAG: // propTag
					hRes = WC_H(MyPropCopyMore(&lpActDst->propTag, &lpActSrc->propTag, MAPIAllocateMore, parent));
					if (FAILED(hRes)) break;
					break;

				case OP_BOUNCE: // scBounceCode
				case OP_DELETE: // union not used
				case OP_MARK_AS_READ:
					break; // Nothing to do!

				default: // error!
				{
					hRes = MAPI_E_INVALID_PARAMETER;
					break;
				}
				}
			}
		}

		if (FAILED(hRes))
		{
			if (!parent) MAPIFreeBuffer(*lppActsDst);
		}

		return hRes;
	}

	// This augmented PropCopyMore is implicitly tied to the built-in MAPIAllocateMore and MAPIAllocateBuffer through
	// the calls to HrCopyRestriction and HrCopyActions. Rewriting those functions to accept function pointers is
	// expensive for no benefit here. So if you borrow this code, be careful if you plan on using other allocators.
	_Check_return_ STDAPI_(SCODE) MyPropCopyMore(
		_In_ LPSPropValue lpSPropValueDest,
		_In_ const _SPropValue* lpSPropValueSrc,
		_In_ ALLOCATEMORE* lpfAllocMore,
		_In_ const VOID* parent)
	{
		auto hRes = S_OK;
		switch (PROP_TYPE(lpSPropValueSrc->ulPropTag))
		{
		case PT_SRESTRICTION:
		case PT_ACTIONS:
		{
			// It's an action or restriction - we know how to copy those:
			memcpy(
				reinterpret_cast<BYTE*>(lpSPropValueDest),
				reinterpret_cast<BYTE*>(const_cast<LPSPropValue>(lpSPropValueSrc)),
				sizeof(SPropValue));
			if (PT_SRESTRICTION == PROP_TYPE(lpSPropValueSrc->ulPropTag))
			{
				LPSRestriction lpNewRes = nullptr;
				hRes = WC_H(HrCopyRestriction(
					reinterpret_cast<LPSRestriction>(lpSPropValueSrc->Value.lpszA), parent, &lpNewRes));
				lpSPropValueDest->Value.lpszA = reinterpret_cast<LPSTR>(lpNewRes);
			}
			else
			{
				ACTIONS* lpNewAct = nullptr;
				hRes = WC_H(HrCopyActions(reinterpret_cast<ACTIONS*>(lpSPropValueSrc->Value.lpszA), parent, &lpNewAct));
				lpSPropValueDest->Value.lpszA = reinterpret_cast<LPSTR>(lpNewAct);
			}
			break;
		}
		default:
			hRes = WC_MAPI(PropCopyMore(
				lpSPropValueDest, const_cast<LPSPropValue>(lpSPropValueSrc), lpfAllocMore, const_cast<LPVOID>(parent)));
		}

		return hRes;
	}

	std::wstring GetProfileName(LPMAPISESSION lpSession)
	{
		LPPROFSECT lpProfSect{};
		std::wstring profileName;

		if (!lpSession) return profileName;

		EC_H_S(lpSession->OpenProfileSection(LPMAPIUID(pbGlobalProfileSectionGuid), nullptr, 0, &lpProfSect));
		if (lpProfSect)
		{
			LPSPropValue lpProfileName{};

			EC_H_S(HrGetOneProp(lpProfSect, PR_PROFILE_NAME_W, &lpProfileName));
			if (lpProfileName && lpProfileName->ulPropTag == PR_PROFILE_NAME_W)
			{
				profileName = std::wstring(lpProfileName->Value.lpszW);
			}

			MAPIFreeBuffer(lpProfileName);

			lpProfSect->Release();
		}

		return profileName;
	}

	bool IsABObject(_In_opt_ LPMAPIPROP lpProp)
	{
		if (!lpProp) return false;

		auto lpPropVal = LPSPropValue{};
		WC_H_S(HrGetOneProp(lpProp, PR_OBJECT_TYPE, &lpPropVal));

		const auto ret = IsABObject(1, lpPropVal);
		MAPIFreeBuffer(lpPropVal);
		return ret;
	}

	bool IsABObject(ULONG ulProps, LPSPropValue lpProps) noexcept
	{
		const auto lpObjectType = PpropFindProp(lpProps, ulProps, PR_OBJECT_TYPE);

		if (lpObjectType && PR_OBJECT_TYPE == lpObjectType->ulPropTag)
		{
			switch (lpObjectType->Value.l)
			{
			case MAPI_ADDRBOOK:
			case MAPI_ABCONT:
			case MAPI_MAILUSER:
			case MAPI_DISTLIST:
				return true;
			}
		}

		return false;
	}
} // namespace mapi