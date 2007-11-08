// MAPIFunctions.h : Stand alone MAPI functions

#pragma once

#include <MapiX.h>
#include <MapiUtil.h>
#include <MAPIform.h>
#include <MSPST.h>

#include <edkmdb.h>
#include <exchform.h>

HRESULT CallOpenEntry(
						LPMDB lpMDB,
						LPADRBOOK lpAB,
						LPMAPICONTAINER lpContainer,
						LPMAPISESSION lpMAPISession,
						LPSBinary lpSBinary,
						LPCIID lpInterface,
						ULONG ulFlags,
						ULONG* ulObjTypeRet,
						LPUNKNOWN* lppUnk);
HRESULT CallOpenEntry(
						LPMDB lpMDB,
						LPADRBOOK lpAB,
						LPMAPICONTAINER lpContainer,
						LPMAPISESSION lpMAPISession,
						ULONG cbEntryID,
						LPENTRYID lpEntryID,
						LPCIID lpInterface,
						ULONG ulFlags,
						ULONG* ulObjTypeRet,
						LPUNKNOWN* lppUnk);
HRESULT	ConcatSPropTagArrays(
							 LPSPropTagArray lpArray1,
							 LPSPropTagArray lpArray2,
							 LPSPropTagArray *lpNewArray);
HRESULT	CopyPropertyAsStream(LPMAPIPROP lpSourcePropObj,LPMAPIPROP lpTargetPropObj, ULONG ulSourceTag, ULONG ulTargetTag);
HRESULT CopyFolderContents(LPMAPIFOLDER lpSrcFolder, LPMAPIFOLDER lpDestFolder,BOOL bCopyAssociatedContents,BOOL bMove,BOOL bSingleCall, HWND hWnd);
HRESULT CopyFolderRules(LPMAPIFOLDER lpSrcFolder, LPMAPIFOLDER lpDestFolder,BOOL bReplace);
HRESULT CopyRestriction(LPSRestriction* lpDestRes, LPSRestriction lpSrcRes, LPVOID lpParent);
HRESULT CopyRestriction(LPSRestriction lpDestRes, LPSRestriction lpSrcRes, LPVOID lpParent);
HRESULT	CopySBinary(LPSBinary psbDest,const LPSBinary psbSrc, LPVOID lpParent);
HRESULT	CopyString(LPTSTR* lpszDestination,LPCTSTR szSource, LPVOID pParent);
HRESULT	CopyStringA(LPSTR* lpszDestination,LPCSTR szSource, LPVOID pParent);
HRESULT	CopyStringW(LPWSTR* lpszDestination,LPCWSTR szSource, LPVOID pParent);
HRESULT CreateNewMailInFolder(
							  LPMAPIFOLDER lpFolder);
HRESULT CreatePropertyStringRestriction(ULONG ulPropTag,
										LPCTSTR szString,
										ULONG ulFuzzyLevel,
										LPVOID lpParent,
										LPSRestriction* lppRes);
HRESULT CreateRangeRestriction(ULONG ulPropTag,
							   LPCTSTR szString,
							   LPVOID lpParent,
							   LPSRestriction* lppRes);
HRESULT DeleteProperty(LPMAPIPROP lpMAPIProp,ULONG ulPropTag);
HRESULT	DeleteToDeletedItems(LPMDB lpMDB, LPMAPIFOLDER lpSourceFolder, LPENTRYLIST lpEIDs, HWND hWnd);
BOOL	FindPropInPropTagArray(LPSPropTagArray lpspTagArray, ULONG ulPropToFind, ULONG* lpulRowFound);
ULONG	GetMAPIObjectType(LPMAPIPROP lpMAPIProp);
HRESULT GetInbox(LPMDB lpMDB, LPMAPIFOLDER* lpInbox);
HRESULT GetParentFolder(LPMAPIFOLDER lpChildFolder, LPMDB lpMDB, LPMAPIFOLDER* lpParentFolder);
HRESULT GetPropsNULL(LPMAPIPROP lpMAPIProp,ULONG ulFlags, ULONG * lpcValues, LPSPropValue *	lppPropArray);
HRESULT GetSpecialFolder(LPMDB lpMDB, ULONG ulFolderPropTag, LPMAPIFOLDER *lpSpecialFolder);
HRESULT IsAttachmentBlocked(LPMAPISESSION lpMAPISession, LPCWSTR pwszFileName, BOOL* pfBlocked);
BOOL	IsDuplicateProp(LPSPropTagArray lpArray, ULONG ulPropTag);
void	MyHexFromBin(LPBYTE lpb, size_t cb, LPTSTR* lpsz);
void	MyBinFromHex(LPCTSTR lpsz, LPBYTE lpb, size_t cb);
HRESULT RemoveOneOff(LPMESSAGE lpMessage, BOOL bRemovePropDef);
HRESULT ResendMessages(
					   LPMAPIFOLDER lpFolder,
					   HWND hWnd);
HRESULT ResendSingleMessage(
							LPMAPIFOLDER lpFolder,
							LPSBinary MessageEID,
							HWND hWnd);
HRESULT ResendSingleMessage(
							LPMAPIFOLDER lpFolder,
							LPMESSAGE lpMessage,
							HWND hWnd);
HRESULT ResetPermissionsOnItems(LPMDB lpMDB, LPMAPIFOLDER lpMAPIFolder);
HRESULT SendTestMessage(
						LPMAPISESSION lpMAPISession,
						LPMAPIFOLDER lpFolder,
						LPCTSTR szRecipient,
						LPCTSTR szBody,
						LPCTSTR szSubject);
HRESULT WrapStreamForRTF(
				 LPSTREAM lpCompressedRTFStream,
				 BOOL bUseWrapEx,
				 ULONG ulFlags,
				 ULONG ulInCodePage,
				 ULONG ulOutCodePage,
				 LPSTREAM FAR * lpUncompressedRTFStream,
				 ULONG FAR * pulStreamFlags);

HRESULT GetNamedPropsByGUID(LPMAPIPROP lpSource, LPGUID lpPropSetGUID, LPSPropTagArray * lpOutArray);
HRESULT CopyNamedProps(LPMAPIPROP lpSource, LPGUID lpPropSetGUID, BOOL bDoMove, BOOL bDoNoReplace, LPMAPIPROP lpTarget, HWND hWnd);

//Unicode support
HRESULT AnsiToUnicode(LPCSTR pszA, LPWSTR* ppszW);
HRESULT UnicodeToAnsi(LPCWSTR pszW, LPSTR* ppszA, size_t cchszW = -1);

BOOL CheckStringProp(LPSPropValue lpProp, ULONG ulPropType);
DWORD ComputeStoreHash(ULONG cbStoreEID, LPENTRYID pbStoreEID, LPCWSTR pwzFileName);
LPWSTR EncodeID(ULONG cbEID, LPENTRYID rgbID);
