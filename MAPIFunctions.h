// MAPIFunctions.h : Stand alone MAPI functions

#pragma once

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
									 _In_opt_ LPSBinary lpSBinary,
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
_Check_return_ HRESULT CopySBinary(_Out_ LPSBinary psbDest, _In_ const LPSBinary psbSrc, _In_ LPVOID lpParent);
_Check_return_ HRESULT CopyStringA(_Deref_out_z_ LPSTR* lpszDestination, _In_z_ LPCSTR szSource, _In_opt_ LPVOID pParent);
_Check_return_ HRESULT CopyStringW(_Deref_out_z_ LPWSTR* lpszDestination, _In_z_ LPCWSTR szSource, _In_opt_ LPVOID pParent);
#ifdef UNICODE
#define CopyString  CopyStringW
#else
#define CopyString  CopyStringA
#endif
_Check_return_ HRESULT CreatePropertyStringRestriction(ULONG ulPropTag,
													   _In_z_ LPCTSTR szString,
													   ULONG ulFuzzyLevel,
													   _In_opt_ LPVOID lpParent,
													   _Deref_out_opt_ LPSRestriction* lppRes);
_Check_return_ HRESULT CreateRangeRestriction(ULONG ulPropTag,
											  _In_z_ LPCTSTR szString,
											  _In_opt_ LPVOID lpParent,
											  _Deref_out_opt_ LPSRestriction* lppRes);
_Check_return_ HRESULT DeleteProperty(_In_ LPMAPIPROP lpMAPIProp,ULONG ulPropTag);
_Check_return_ HRESULT DeleteToDeletedItems(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpSourceFolder, _In_ LPENTRYLIST lpEIDs, _In_ HWND hWnd);
_Check_return_ bool    FindPropInPropTagArray(_In_ LPSPropTagArray lpspTagArray, ULONG ulPropToFind, _Out_ ULONG* lpulRowFound);
_Check_return_ ULONG   GetMAPIObjectType(_In_opt_ LPMAPIPROP lpMAPIProp);
_Check_return_ HRESULT GetParentFolder(_In_ LPMAPIFOLDER lpChildFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER* lpParentFolder);
_Check_return_ HRESULT GetPropsNULL(_In_ LPMAPIPROP lpMAPIProp, ULONG ulFlags, _Out_ ULONG* lpcValues, _Deref_out_opt_ LPSPropValue* lppPropArray);
_Check_return_ HRESULT GetSpecialFolderEID(_In_ LPMDB lpMDB, ULONG ulFolderPropTag, _Out_opt_ ULONG* lpcbeid, _Deref_out_opt_ LPENTRYID* lppeid);
_Check_return_ HRESULT IsAttachmentBlocked(_In_ LPMAPISESSION lpMAPISession, _In_z_ LPCWSTR pwszFileName, _Out_ bool* pfBlocked);
_Check_return_ bool    IsDuplicateProp(_In_ LPSPropTagArray lpArray, ULONG ulPropTag);
_Check_return_ HRESULT ManuallyEmptyFolder(_In_ LPMAPIFOLDER lpFolder, BOOL bAssoc, BOOL bHardDelete);
void    MyHexFromBin(_In_opt_count_(cb) LPBYTE lpb, size_t cb, bool bPrependCB, _Deref_out_opt_z_ LPTSTR* lpsz);
_Check_return_ bool MyBinFromHex(_In_z_ LPCTSTR lpsz, _Out_opt_bytecapcount_(*lpcb) LPBYTE lpb, _Inout_ ULONG* lpcb);
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
									   _In_z_ LPCTSTR szRecipient,
									   _In_z_ LPCTSTR szBody,
									   _In_z_ LPCTSTR szSubject,
									   _In_z_ LPCTSTR szClass);
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

_Check_return_ bool CheckStringProp(_In_opt_ LPSPropValue lpProp, ULONG ulPropType);
_Check_return_ DWORD ComputeStoreHash(ULONG cbStoreEID, _In_count_(cbStoreEID) LPBYTE pbStoreEID, _In_opt_z_ LPCSTR pszFileName, _In_opt_z_ LPCWSTR pwzFileName, bool bPublicStore);
_Check_return_ LPWSTR EncodeID(ULONG cbEID, _In_ LPENTRYID rgbID);
_Check_return_ LPWSTR DecodeID(ULONG cbBuffer, _In_count_(cbBuffer) LPBYTE lpbBuffer);

HRESULT HrEmsmdbUIDFromStore(_In_ LPMAPISESSION pmsess,
							 _In_ MAPIUID const * const puidService,
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
static LPSTR FolderNames[] = {
	"",
	"Calendar",
	"Contacts",
	"Journal",
	"Notes",
	"Tasks",
	"Reminders",
	"Drafts",
	"Sent Items",
	"Outbox",
	"Deleted Items",
	"Finder",
	"IPM_SUBTREE",
	"Inbox",
	"Local Freebusy",
	"Conflicts",
	"Sync Issues",
	"Local Failures",
	"Server Failures",
	"Junk E-mail",
};

STDMETHODIMP OpenDefaultFolder(_In_ ULONG ulFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER *lpFolder);
STDMETHODIMP GetDefaultFolderEID(
	_In_ ULONG ulFolder,
	_In_ LPMDB lpMDB,
	_Out_opt_ ULONG* lpcbeid,
	_Deref_out_opt_ LPENTRYID* lppeid);

CString GetTitle(LPMAPIPROP lpMAPIProp);
bool UnwrapContactEntryID(_In_ ULONG cbIn, _In_ LPBYTE lpbIn, _Out_ ULONG* lpcbOut, _Out_ LPBYTE* lppbOut);

LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB);
HRESULT CopyTo(HWND hWnd, _In_ LPMAPIPROP lpSource, _In_ LPMAPIPROP lpDest, LPCGUID lpGUID, _In_opt_ LPSPropTagArray lpTagArray, bool bIsAB, bool bAllowUI);