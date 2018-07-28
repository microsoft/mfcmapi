// Stand alone MAPI functions
#pragma once
#include "Interpret/GUIDArray.h"

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
		else ASSERT(false);
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

	_Check_return_ LPSPropTagArray
	ConcatSPropTagArrays(_In_ LPSPropTagArray lpArray1, _In_opt_ LPSPropTagArray lpArray2);
	_Check_return_ HRESULT CopyPropertyAsStream(
		_In_ LPMAPIPROP lpSourcePropObj,
		_In_ LPMAPIPROP lpTargetPropObj,
		ULONG ulSourceTag,
		ULONG ulTargetTag);
	_Check_return_ void CopyFolderContents(
		_In_ LPMAPIFOLDER lpSrcFolder,
		_In_ LPMAPIFOLDER lpDestFolder,
		bool bCopyAssociatedContents,
		bool bMove,
		bool bSingleCall,
		_In_ HWND hWnd);
	_Check_return_ HRESULT
	CopyFolderRules(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bReplace);
	_Check_return_ SBinary CopySBinary(_In_ const _SBinary& src, _In_ LPVOID parent = nullptr);
	_Check_return_ LPSBinary CopySBinary(_In_ const _SBinary* src);
	_Check_return_ LPSTR CopyStringA(_In_z_ LPCSTR szSource, _In_opt_ LPVOID pParent);
	_Check_return_ LPWSTR CopyStringW(_In_z_ LPCWSTR szSource, _In_opt_ LPVOID pParent);
#ifdef UNICODE
#define CopyString CopyStringW
#else
#define CopyString CopyStringA
#endif
	_Check_return_ LPSRestriction CreatePropertyStringRestriction(
		ULONG ulPropTag,
		_In_ const std::wstring& szString,
		ULONG ulFuzzyLevel,
		_In_opt_ LPVOID parent);
	_Check_return_ LPSRestriction
	CreateRangeRestriction(ULONG ulPropTag, _In_ const std::wstring& szString, _In_opt_ LPVOID);
	_Check_return_ HRESULT DeleteProperty(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);
	_Check_return_ HRESULT
	DeleteToDeletedItems(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpSourceFolder, _In_ LPENTRYLIST lpEIDs, _In_ HWND hWnd);
	_Check_return_ bool
	FindPropInPropTagArray(_In_ LPSPropTagArray lpspTagArray, ULONG ulPropToFind, _Out_ ULONG* lpulRowFound);
	_Check_return_ ULONG GetMAPIObjectType(_In_opt_ LPMAPIPROP lpMAPIProp);
	_Check_return_ LPMAPIFOLDER GetParentFolder(_In_ LPMAPIFOLDER lpChildFolder, _In_ LPMDB lpMDB);
	_Check_return_ HRESULT GetPropsNULL(
		_In_ LPMAPIPROP lpMAPIProp,
		ULONG ulFlags,
		_Out_ ULONG* lpcValues,
		_Deref_out_opt_ LPSPropValue* lppPropArray);
	_Check_return_ HRESULT GetSpecialFolderEID(
		_In_ LPMDB lpMDB,
		ULONG ulFolderPropTag,
		_Out_opt_ ULONG* lpcbeid,
		_Deref_out_opt_ LPENTRYID* lppeid);
	_Check_return_ HRESULT
	IsAttachmentBlocked(_In_ LPMAPISESSION lpMAPISession, _In_z_ LPCWSTR pwszFileName, _Out_ bool* pfBlocked);
	_Check_return_ bool IsDuplicateProp(_In_ LPSPropTagArray lpArray, ULONG ulPropTag);
	_Check_return_ HRESULT ManuallyEmptyFolder(_In_ LPMAPIFOLDER lpFolder, BOOL bAssoc, BOOL bHardDelete);
	_Check_return_ LPBYTE ByteVectorToMAPI(const std::vector<BYTE>& bin, LPVOID lpParent);
	_Check_return_ HRESULT RemoveOneOff(_In_ LPMESSAGE lpMessage, bool bRemovePropDef);
	_Check_return_ HRESULT ResendMessages(_In_ LPMAPIFOLDER lpFolder, _In_ HWND hWnd);
	_Check_return_ HRESULT ResendSingleMessage(_In_ LPMAPIFOLDER lpFolder, _In_ LPSBinary MessageEID, _In_ HWND hWnd);
	_Check_return_ HRESULT ResendSingleMessage(_In_ LPMAPIFOLDER lpFolder, _In_ LPMESSAGE lpMessage, _In_ HWND hWnd);
	_Check_return_ HRESULT ResetPermissionsOnItems(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpMAPIFolder);
	_Check_return_ HRESULT SendTestMessage(
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMAPIFOLDER lpFolder,
		_In_ const std::wstring& szRecipient,
		_In_ const std::wstring& szBody,
		_In_ const std::wstring& szSubject,
		_In_ const std::wstring& szClass);
	_Check_return_ HRESULT WrapStreamForRTF(
		_In_ LPSTREAM lpCompressedRTFStream,
		bool bUseWrapEx,
		ULONG ulFlags,
		ULONG ulInCodePage,
		ULONG ulOutCodePage,
		_Deref_out_ LPSTREAM* lpUncompressedRTFStream,
		_Out_opt_ ULONG* pulStreamFlags);

	_Check_return_ LPSPropTagArray GetNamedPropsByGUID(_In_ LPMAPIPROP lpSource, _In_ LPGUID lpPropSetGUID);
	_Check_return_ HRESULT CopyNamedProps(
		_In_ LPMAPIPROP lpSource,
		_In_ LPGUID lpPropSetGUID,
		bool bDoMove,
		bool bDoNoReplace,
		_In_ LPMAPIPROP lpTarget,
		_In_ HWND hWnd);

	_Check_return_ bool CheckStringProp(_In_opt_ const _SPropValue* lpProp, ULONG ulPropType);
	_Check_return_ DWORD ComputeStoreHash(
		ULONG cbStoreEID,
		_In_count_(cbStoreEID) LPBYTE pbStoreEID,
		_In_opt_z_ LPCSTR pszFileName,
		_In_opt_z_ LPCWSTR pwzFileName,
		bool bPublicStore);
	_Check_return_ std::wstring EncodeID(ULONG cbEID, _In_ LPENTRYID rgbID);
	_Check_return_ std::wstring DecodeID(ULONG cbBuffer, _In_count_(cbBuffer) LPBYTE lpbBuffer);

	HRESULT
	HrEmsmdbUIDFromStore(_In_ LPMAPISESSION pmsess, _In_ const MAPIUID* puidService, _Out_opt_ MAPIUID* pEmsmdbUID);
	bool FExchangePrivateStore(_In_ LPMAPIUID lpmapiuid);
	bool FExchangePublicStore(_In_ LPMAPIUID lpmapiuid);

	enum
	{
		DEFAULT_UNSPECIFIED,
		DEFAULT_CALENDAR,
		DEFAULT_CONTACTS,
		DEFAULT_JOURNAL,
		DEFAULT_NOTES,
		DEFAULT_TASKS,
		DEFAULT_REMINDERS,
		DEFAULT_DRAFTS,
		DEFAULT_SENTITEMS,
		DEFAULT_OUTBOX,
		DEFAULT_DELETEDITEMS,
		DEFAULT_FINDER,
		DEFAULT_IPM_SUBTREE,
		DEFAULT_INBOX,
		DEFAULT_LOCALFREEBUSY,
		DEFAULT_CONFLICTS,
		DEFAULT_SYNCISSUES,
		DEFAULT_LOCALFAILURES,
		DEFAULT_SERVERFAILURES,
		DEFAULT_JUNKMAIL,
		NUM_DEFAULT_PROPS
	};

	// Keep this in sync with the NUM_DEFAULT_PROPS enum above
	static LPCWSTR FolderNames[] = {
		L"",
		L"Calendar",
		L"Contacts",
		L"Journal",
		L"Notes",
		L"Tasks",
		L"Reminders",
		L"Drafts",
		L"Sent Items",
		L"Outbox",
		L"Deleted Items",
		L"Finder",
		L"IPM_SUBTREE",
		L"Inbox",
		L"Local Freebusy",
		L"Conflicts",
		L"Sync Issues",
		L"Local Failures",
		L"Server Failures",
		L"Junk E-mail",
	};

	LPMAPIFOLDER OpenDefaultFolder(_In_ ULONG ulFolder, _In_ LPMDB lpMDB);
	LPSBinary GetDefaultFolderEID(_In_ ULONG ulFolder, _In_ LPMDB lpMDB);

	std::wstring GetTitle(LPMAPIPROP lpMAPIProp);
	bool UnwrapContactEntryID(_In_ ULONG cbIn, _In_ LPBYTE lpbIn, _Out_ ULONG* lpcbOut, _Out_ LPBYTE* lppbOut);

	LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB);
	HRESULT CopyTo(
		HWND hWnd,
		_In_ LPMAPIPROP lpSource,
		_In_ LPMAPIPROP lpDest,
		LPCGUID lpGUID,
		_In_opt_ LPSPropTagArray lpTagArray,
		bool bIsAB,
		bool bAllowUI);

	// Augemented version of HrGetOneProp which allows passing flags to underlying GetProps
	// Useful for passing fMapiUnicode for unspecified string/stream types
	HRESULT
	HrGetOnePropEx(_In_ LPMAPIPROP lpMAPIProp, _In_ ULONG ulPropTag, _In_ ULONG ulFlags, _Out_ LPSPropValue* lppProp);

	void ForceRop(_In_ LPMDB lpMDB);

	_Check_return_ LPSPropValue GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);
	_Check_return_ LPSPropValue GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);

	_Check_return_ STDAPI HrCopyRestriction(
		_In_ const _SRestriction* lpResSrc, // source restriction ptr
		_In_opt_ const LPVOID lpObject, // ptr to existing MAPI buffer
		_In_ LPSRestriction* lppResDest // dest restriction buffer ptr
	);

	_Check_return_ HRESULT HrCopyRestrictionArray(
		_In_ const _SRestriction* lpResSrc, // source restriction
		_In_ const LPVOID lpObject, // ptr to existing MAPI buffer
		ULONG cRes, // # elements in array
		_In_count_(cRes) LPSRestriction lpResDest // destination restriction
	);

	_Check_return_ STDAPI_(SCODE) MyPropCopyMore(
		_In_ LPSPropValue lpSPropValueDest,
		_In_ const _SPropValue* lpSPropValueSrc,
		_In_ ALLOCATEMORE* lpfAllocMore,
		_In_ LPVOID lpvObject);

	_Check_return_ STDMETHODIMP MyOpenStreamOnFile(
		_In_ LPALLOCATEBUFFER lpAllocateBuffer,
		_In_ LPFREEBUFFER lpFreeBuffer,
		ULONG ulFlags,
		_In_ const std::wstring& lpszFileName,
		_Out_ LPSTREAM* lppStream);
}