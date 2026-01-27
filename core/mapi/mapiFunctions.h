// Stand alone MAPI functions
#pragma once

namespace mapi
{
	// http://msdn.microsoft.com/en-us/library/office/dn433223.aspx
#pragma pack(4)
	struct CONTAB_ENTRYID
	{
		BYTE abFlags[4];
		MAPIUID muid;
		ULONG ulVersion;
		ULONG ulType;
		ULONG ulIndex;
		ULONG cbeid;
		BYTE abeid[1];
	};
	typedef CONTAB_ENTRYID* LPCONTAB_ENTRYID;
#pragma pack()

	// http://msdn.microsoft.com/en-us/library/office/dn433221.aspx
#pragma pack(4)
	struct DIR_ENTRYID
	{
		BYTE abFlags[4];
		MAPIUID muid;
		ULONG ulVersion;
		ULONG ulType;
		MAPIUID muidID;
	};
	typedef DIR_ENTRYID* LPDIR_ENTRYID;
#pragma pack()

#define CbNewROWLIST(_centries) (offsetof(ROWLIST, aEntries) + (_centries) * sizeof(ROWENTRY))
#define MAXNewROWLIST ((ULONG_MAX - offsetof(ROWLIST, aEntries)) / sizeof(ROWENTRY))
#define MAXMessageClassArray ((ULONG_MAX - offsetof(SMessageClassArray, aMessageClass)) / sizeof(LPCSTR))
#define MAXNewADRLIST ((ULONG_MAX - offsetof(ADRLIST, aEntries)) / sizeof(ADRENTRY))

	typedef ACTIONS* LPACTIONS;

	struct CopyDetails
	{
		bool valid{};
		ULONG flags{};
		GUID guid{};
		LPMAPIPROGRESS progress{};
		ULONG_PTR uiParam{};
		LPSPropTagArray excludedTags{};
		bool allocated{};
		void clean() const noexcept
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
		const ULONG cbEntryID,
		const _In_opt_ LPENTRYID lpEntryID,
		_In_opt_ LPCIID lpInterface,
		ULONG ulFlags,
		_Out_opt_ ULONG* ulObjTypeRet); // optional - can be NULL

	template <class T>
	T CallOpenEntry(
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

	_Check_return_ SBinary CopySBinary(_In_ const _SBinary& src, _In_opt_ const VOID* parent = nullptr);
	_Check_return_ LPSBinary CopySBinary(_In_ const _SBinary* src);

	_Check_return_ LPSTR CopyStringA(_In_z_ LPCSTR src, _In_opt_ const VOID* parent);
	_Check_return_ LPWSTR CopyStringW(_In_z_ LPCWSTR src, _In_opt_ const VOID* parent);
#ifdef UNICODE
#define CopyString CopyStringW
#else
#define CopyString CopyStringA
#endif

	_Check_return_ LPSPropTagArray
	ConcatSPropTagArrays(_In_ LPSPropTagArray lpArray1, _In_opt_ LPSPropTagArray lpArray2);
	_Check_return_ HRESULT CopyPropertyAsStream(
		_In_ LPMAPIPROP lpSourcePropObj,
		_In_ LPMAPIPROP lpTargetPropObj,
		ULONG ulSourceTag,
		ULONG ulTargetTag);
	void CopyFolderContents(
		_In_ LPMAPIFOLDER lpSrcFolder,
		_In_ LPMAPIFOLDER lpDestFolder,
		bool bCopyAssociatedContents,
		bool bMove,
		bool bSingleCall,
		_In_ HWND hWnd);
	_Check_return_ HRESULT
	CopyFolderRules(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bReplace);
	_Check_return_ LPSRestriction CreatePropertyStringRestriction(
		ULONG ulPropTag,
		_In_ const std::wstring& szString,
		ULONG ulFuzzyLevel,
		_In_opt_ const VOID* parent);
	_Check_return_ LPSRestriction
	CreateRangeRestriction(ULONG ulPropTag, _In_ const std::wstring& szString, _In_opt_ const VOID* parent);
	_Check_return_ HRESULT DeleteProperty(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);
	_Check_return_ HRESULT
	DeleteToDeletedItems(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpSourceFolder, _In_ LPENTRYLIST lpEIDs, _In_ HWND hWnd);
	_Check_return_ bool
	FindPropInPropTagArray(_In_ LPSPropTagArray lpspTagArray, ULONG ulPropToFind, _Out_ ULONG* lpulRowFound) noexcept;
	_Check_return_ LPMAPIFOLDER GetParentFolder(_In_ LPMAPIFOLDER lpChildFolder, _In_ LPMDB lpMDB);
	_Check_return_ LPSBinary GetSpecialFolderEID(_In_ LPMDB lpMDB, ULONG ulFolderPropTag);
	_Check_return_ HRESULT
	IsAttachmentBlocked(_In_ LPMAPISESSION lpMAPISession, _In_z_ LPCWSTR pwszFileName, _Out_ bool* pfBlocked);
	_Check_return_ bool IsDuplicateProp(_In_ LPSPropTagArray lpArray, ULONG ulPropTag) noexcept;
	_Check_return_ HRESULT ManuallyEmptyFolder(_In_ LPMAPIFOLDER lpFolder, BOOL bAssoc, BOOL bHardDelete);
	_Check_return_ LPBYTE ByteVectorToMAPI(const std::vector<BYTE>& bin, const VOID* parent);
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

	_Check_return_ DWORD ComputeStoreHash(
		ULONG cbStoreEID,
		_In_count_(cbStoreEID) LPBYTE pbStoreEID,
		_In_opt_z_ LPCSTR pszFileName,
		_In_opt_z_ LPCWSTR pwzFileName,
		bool bPublicStore);

	HRESULT
	HrEmsmdbUIDFromStore(_In_ LPMAPISESSION pmsess, _In_ const MAPIUID* puidService, _Out_opt_ MAPIUID* pEmsmdbUID);
	bool FExchangePrivateStore(_In_ LPMAPIUID lpmapiuid) noexcept;
	bool FExchangePublicStore(_In_ LPMAPIUID lpmapiuid) noexcept;

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
	bool UnwrapContactEntryID(_In_ ULONG cbIn, _In_ LPBYTE lpbIn, _Out_ ULONG* lpcbOut, _Out_ LPBYTE* lppbOut) noexcept;

	LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB);

	// Augemented version of HrGetOneProp which allows passing flags to underlying GetProps
	// Useful for passing fMapiUnicode for unspecified string/stream types
	HRESULT
	HrGetOnePropEx(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ ULONG ulPropTag,
		_In_ ULONG ulFlags,
		_Out_ LPSPropValue* lppProp) noexcept;

	void ForceRop(_In_ LPMDB lpMDB);

	_Check_return_ HRESULT
	HrDupPropset(int cprop, _In_count_(cprop) LPSPropValue rgprop, _In_ const VOID* parent, _In_ LPSPropValue* prgprop);

	_Check_return_ STDAPI HrCopyRestriction(
		_In_ const _SRestriction* lpResSrc, // source restriction ptr
		_In_opt_ const VOID* lpObject, // ptr to existing MAPI buffer
		_In_ LPSRestriction* lppResDest // dest restriction buffer ptr
	);

	_Check_return_ HRESULT HrCopyRestrictionArray(
		_In_ const _SRestriction* lpResSrc, // source restriction
		_In_ const VOID* lpObject, // ptr to existing MAPI buffer
		ULONG cRes, // # elements in array
		_In_count_(cRes) LPSRestriction lpResDest // destination restriction
	);

	_Check_return_ STDAPI_(SCODE) MyPropCopyMore(
		_In_ LPSPropValue lpSPropValueDest,
		_In_ const _SPropValue* lpSPropValueSrc,
		_In_ ALLOCATEMORE* lpfAllocMore,
		_In_ const VOID* parent);

#pragma warning(push)
#pragma warning(disable : 26482) // Warning C26482 Only index into arrays using constant expressions (bounds.2).
	// Get a tag from a const array
	inline ULONG getTag(const SPropTagArray* tag, ULONG i) noexcept { return tag->aulPropTag[i]; }
	inline ULONG getTag(const SPropTagArray& tag, ULONG i) noexcept { return tag.aulPropTag[i]; }
	// Get a reference to a tag usable for setting.
	// Cannot be used with const arrays.
	inline ULONG& setTag(SPropTagArray* tag, ULONG i) noexcept { return tag->aulPropTag[i]; }
	inline ULONG& setTag(SPropTagArray& tag, ULONG i) noexcept { return tag.aulPropTag[i]; }
#pragma warning(pop)

#pragma warning(push)
#pragma warning( \
	disable : 26476) // Warning C26476 Expression/symbol '' uses a naked union '' with multiple type pointers: Use variant instead (type.7).
	inline const SBinary& getBin(_In_ const _SPropValue* prop) noexcept
	{
		if (!prop) assert(false);

		return prop->Value.bin;
	}
	inline SBinary& setBin(_In_ _SPropValue* prop) noexcept
	{
		if (!prop) assert(false);

		return prop->Value.bin;
	}
	inline const SBinary& getBin(_In_ const _SPropValue& prop) noexcept { return prop.Value.bin; }
	inline SBinary& setBin(_In_ _SPropValue& prop) noexcept { return prop.Value.bin; }
#pragma warning(pop)

	inline LPENTRYID toEntryID(const std::vector<BYTE>& eid) noexcept
	{
		return reinterpret_cast<LPENTRYID>(const_cast<BYTE*>(eid.data()));
	}

	std::wstring GetProfileName(LPMAPISESSION lpSession);

	bool IsABObject(_In_opt_ LPMAPIPROP lpProp);
	bool IsABObject(ULONG ulProps, LPSPropValue lpProps) noexcept;

	inline _Check_return_ LPSPropValue
	FindProp(_In_ const SPropValue* lpPropArray, _In_ const ULONG cValues, _In_ const ULONG ulPropTag)
	{
		return PpropFindProp(const_cast<SPropValue*>(lpPropArray), cValues, ulPropTag);
	}

	inline _Check_return_ HRESULT DupeProps(
		_In_ const int cValues,
		_In_ const SPropValue* lpPropArray,
		_In_ const LPALLOCATEBUFFER lpAllocateBuffer,
		_Out_ LPSPropValue* lppPropArray)
	{
		return ScDupPropset(cValues, const_cast<LPSPropValue>(lpPropArray), lpAllocateBuffer, lppPropArray);
	}
} // namespace mapi