// Stand alone MAPI functions
#pragma once
#include "Interpret/GUIDArray.h"

namespace mapi
{
	// Safely cast across MAPI interfaces. Result is addrefed and must be released.
	_Check_return_ template <class T> T safe_cast(IUnknown* src)
	{
		if (!src) return nullptr;
		auto hRes = S_OK;
		T ret = nullptr;

		auto iid = IID_IUnknown;
		if (std::is_same<T, LPMAPIFOLDER>::value)
		{
			iid = IID_IMAPIFolder;
		}
		else if (std::is_same<T, LPMAPICONTAINER>::value)
		{
			iid = IID_IMAPIContainer;
		}
		else if (std::is_same<T, LPMAILUSER>::value)
		{
			iid = IID_IMailUser;
		}
		else if (std::is_same<T, LPABCONT>::value)
		{
			iid = IID_IABContainer;
		}
		else if (std::is_same<T, LPMESSAGE>::value)
		{
			iid = IID_IMessage;
		}
		else if (std::is_same<T, LPMDB>::value)
		{
			iid = IID_IMsgStore;
		}
		else if (std::is_same<T, LPMAPIFORMINFO>::value)
		{
			iid = IID_IMAPIFormInfo;
		}
		else if (std::is_same<T, IAttachmentSecurity*>::value)
		{
			iid = guid::IID_IAttachmentSecurity;
		}
		else if (std::is_same<T, LPSERVICEADMIN2>::value)
		{
			iid = IID_IMsgServiceAdmin2;
		}
		else if (std::is_same<T, LPEXCHANGEMANAGESTORE>::value)
		{
			iid = IID_IExchangeManageStore;
		}
		else if (std::is_same<T, LPEXCHANGEMANAGESTORE3>::value)
		{
			iid = IID_IExchangeManageStore3;
		}
		else if (std::is_same<T, LPEXCHANGEMANAGESTORE4>::value)
		{
			iid = IID_IExchangeManageStore4;
		}
		else if (std::is_same<T, LPEXCHANGEMANAGESTORE5>::value)
		{
			iid = guid::IID_IExchangeManageStore5;
		}
		else if (std::is_same<T, LPEXCHANGEMANAGESTOREEX>::value)
		{
			iid = guid::IID_IExchangeManageStoreEx;
		}
		else
		{
			ASSERT(false);
		}

		WC_H(src->QueryInterface(iid, reinterpret_cast<LPVOID*>(&ret)));
		output::DebugPrint(DBGGeneric, L"safe_cast: iid =%ws, src = %p, ret = %p\n",
			guid::GUIDToStringAndName(&iid).c_str(),
			src,
			ret);

		return ret;
	}

	_Check_return_ HRESULT CallOpenEntry(
		_In_opt_ LPMDB lpMDB,
		_In_opt_ LPADRBOOK lpAB,
		_In_opt_ LPMAPICONTAINER lpContainer,
		_In_opt_ LPMAPISESSION lpMAPISession,
		ULONG cbEntryID,
		_In_opt_ LPENTRYID lpEntryID,
		_In_opt_ LPCIID lpInterface,
		ULONG ulFlags,
		_Out_opt_ ULONG* ulObjTypeRet,
		_Deref_out_opt_ LPUNKNOWN* lppUnk);
	_Check_return_ HRESULT CallOpenEntry(
		_In_opt_ LPMDB lpMDB,
		_In_opt_ LPADRBOOK lpAB,
		_In_opt_ LPMAPICONTAINER lpContainer,
		_In_opt_ LPMAPISESSION lpMAPISession,
		_In_opt_ const SBinary* lpSBinary,
		_In_opt_ LPCIID lpInterface,
		ULONG ulFlags,
		_Out_opt_ ULONG* ulObjTypeRet,
		_Deref_out_opt_ LPUNKNOWN* lppUnk);
	_Check_return_ HRESULT ConcatSPropTagArrays(
		_In_ LPSPropTagArray lpArray1,
		_In_opt_ LPSPropTagArray lpArray2,
		_Deref_out_opt_ LPSPropTagArray *lpNewArray);
	_Check_return_ HRESULT CopyPropertyAsStream(_In_ LPMAPIPROP lpSourcePropObj, _In_ LPMAPIPROP lpTargetPropObj, ULONG ulSourceTag, ULONG ulTargetTag);
	_Check_return_ HRESULT CopyFolderContents(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bCopyAssociatedContents, bool bMove, bool bSingleCall, _In_ HWND hWnd);
	_Check_return_ HRESULT CopyFolderRules(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bReplace);
	_Check_return_ HRESULT CopySBinary(_Out_ LPSBinary psbDest, _In_ const _SBinary* psbSrc, _In_ LPVOID lpParent);
	_Check_return_ HRESULT CopyStringA(_Deref_out_z_ LPSTR* lpszDestination, _In_z_ LPCSTR szSource, _In_opt_ LPVOID pParent);
	_Check_return_ HRESULT CopyStringW(_Deref_out_z_ LPWSTR* lpszDestination, _In_z_ LPCWSTR szSource, _In_opt_ LPVOID pParent);
#ifdef UNICODE
#define CopyString CopyStringW
#else
#define CopyString CopyStringA
#endif
	_Check_return_ HRESULT CreatePropertyStringRestriction(ULONG ulPropTag,
		_In_ const std::wstring& szString,
		ULONG ulFuzzyLevel,
		_In_opt_ LPVOID lpParent,
		_Deref_out_opt_ LPSRestriction* lppRes);
	_Check_return_ HRESULT CreateRangeRestriction(ULONG ulPropTag,
		_In_ const std::wstring& szString,
		_In_opt_ LPVOID lpParent,
		_Deref_out_opt_ LPSRestriction* lppRes);
	_Check_return_ HRESULT DeleteProperty(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);
	_Check_return_ HRESULT DeleteToDeletedItems(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpSourceFolder, _In_ LPENTRYLIST lpEIDs, _In_ HWND hWnd);
	_Check_return_ bool FindPropInPropTagArray(_In_ LPSPropTagArray lpspTagArray, ULONG ulPropToFind, _Out_ ULONG* lpulRowFound);
	_Check_return_ ULONG GetMAPIObjectType(_In_opt_ LPMAPIPROP lpMAPIProp);
	_Check_return_ HRESULT GetParentFolder(_In_ LPMAPIFOLDER lpChildFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER* lpParentFolder);
	_Check_return_ HRESULT GetPropsNULL(_In_ LPMAPIPROP lpMAPIProp, ULONG ulFlags, _Out_ ULONG* lpcValues, _Deref_out_opt_ LPSPropValue* lppPropArray);
	_Check_return_ HRESULT GetSpecialFolderEID(_In_ LPMDB lpMDB, ULONG ulFolderPropTag, _Out_opt_ ULONG* lpcbeid, _Deref_out_opt_ LPENTRYID* lppeid);
	_Check_return_ HRESULT IsAttachmentBlocked(_In_ LPMAPISESSION lpMAPISession, _In_z_ LPCWSTR pwszFileName, _Out_ bool* pfBlocked);
	_Check_return_ bool IsDuplicateProp(_In_ LPSPropTagArray lpArray, ULONG ulPropTag);
	_Check_return_ HRESULT ManuallyEmptyFolder(_In_ LPMAPIFOLDER lpFolder, BOOL bAssoc, BOOL bHardDelete);
	_Check_return_ LPBYTE ByteVectorToMAPI(const std::vector<BYTE>& bin, LPVOID lpParent);
	_Check_return_ HRESULT RemoveOneOff(_In_ LPMESSAGE lpMessage, bool bRemovePropDef);
	_Check_return_ HRESULT ResendMessages(
		_In_ LPMAPIFOLDER lpFolder,
		_In_ HWND hWnd);
	_Check_return_ HRESULT ResendSingleMessage(
		_In_ LPMAPIFOLDER lpFolder,
		_In_ LPSBinary MessageEID,
		_In_ HWND hWnd);
	_Check_return_ HRESULT ResendSingleMessage(
		_In_ LPMAPIFOLDER lpFolder,
		_In_ LPMESSAGE lpMessage,
		_In_ HWND hWnd);
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

	_Check_return_ HRESULT GetNamedPropsByGUID(_In_ LPMAPIPROP lpSource, _In_ LPGUID lpPropSetGUID, _Deref_out_ LPSPropTagArray * lpOutArray);
	_Check_return_ HRESULT CopyNamedProps(_In_ LPMAPIPROP lpSource, _In_ LPGUID lpPropSetGUID, bool bDoMove, bool bDoNoReplace, _In_ LPMAPIPROP lpTarget, _In_ HWND hWnd);

	_Check_return_ bool CheckStringProp(_In_opt_ const _SPropValue* lpProp, ULONG ulPropType);
	_Check_return_ DWORD ComputeStoreHash(ULONG cbStoreEID, _In_count_(cbStoreEID) LPBYTE pbStoreEID, _In_opt_z_ LPCSTR pszFileName, _In_opt_z_ LPCWSTR pwzFileName, bool bPublicStore);
	_Check_return_ std::wstring EncodeID(ULONG cbEID, _In_ LPENTRYID rgbID);
	_Check_return_ std::wstring DecodeID(ULONG cbBuffer, _In_count_(cbBuffer) LPBYTE lpbBuffer);

	HRESULT HrEmsmdbUIDFromStore(_In_ LPMAPISESSION pmsess,
		_In_ const MAPIUID* puidService,
		_Out_opt_ MAPIUID* pEmsmdbUID);
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

	STDMETHODIMP OpenDefaultFolder(_In_ ULONG ulFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER *lpFolder);
	STDMETHODIMP GetDefaultFolderEID(
		_In_ ULONG ulFolder,
		_In_ LPMDB lpMDB,
		_Out_opt_ ULONG* lpcbeid,
		_Deref_out_opt_ LPENTRYID* lppeid);

	std::wstring GetTitle(LPMAPIPROP lpMAPIProp);
	bool UnwrapContactEntryID(_In_ ULONG cbIn, _In_ LPBYTE lpbIn, _Out_ ULONG* lpcbOut, _Out_ LPBYTE* lppbOut);

	LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB);
	HRESULT CopyTo(HWND hWnd, _In_ LPMAPIPROP lpSource, _In_ LPMAPIPROP lpDest, LPCGUID lpGUID, _In_opt_ LPSPropTagArray lpTagArray, bool bIsAB, bool bAllowUI);

	// Augemented version of HrGetOneProp which allows passing flags to underlying GetProps
	// Useful for passing fMapiUnicode for unspecified string/stream types
	HRESULT HrGetOnePropEx(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ ULONG ulPropTag,
		_In_ ULONG ulFlags,
		_Out_ LPSPropValue* lppProp);

	void ForceRop(_In_ LPMDB lpMDB);

	_Check_return_ HRESULT GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);
	_Check_return_ HRESULT GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);

	_Check_return_ STDAPI HrCopyRestriction(
		_In_ const _SRestriction* lpResSrc, // source restriction ptr
		_In_opt_ LPVOID lpObject, // ptr to existing MAPI buffer
		_In_ LPSRestriction* lppResDest // dest restriction buffer ptr
	);

	_Check_return_ HRESULT HrCopyRestrictionArray(
		_In_ const _SRestriction* lpResSrc, // source restriction
		_In_ LPVOID lpObject, // ptr to existing MAPI buffer
		ULONG cRes, // # elements in array
		_In_count_(cRes) LPSRestriction lpResDest // destination restriction
	);

	_Check_return_ STDAPI_(SCODE) MyPropCopyMore(_In_ LPSPropValue lpSPropValueDest,
		_In_ const _SPropValue* lpSPropValueSrc,
		_In_ ALLOCATEMORE * lpfAllocMore,
		_In_ LPVOID lpvObject);

	_Check_return_ STDMETHODIMP MyOpenStreamOnFile(_In_ LPALLOCATEBUFFER lpAllocateBuffer,
		_In_ LPFREEBUFFER lpFreeBuffer,
		ULONG ulFlags,
		_In_ const std::wstring& lpszFileName,
		_Out_ LPSTREAM* lppStream);
}