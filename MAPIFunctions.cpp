// MAPIfunctions.cpp : Collection of useful MAPI functions

#include "stdafx.h"
#include "MAPIFunctions.h"
#include "String.h"
#include "MAPIABFunctions.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "ImportProcs.h"
#include "ExtraPropTags.h"
#include "MAPIProgress.h"
#include "Guids.h"
#include "NamedPropCache.h"
#include "SmartView/SmartView.h"
#include <Dialogs/Editors/TagArrayEditor.h>

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

_Check_return_ HRESULT CallOpenEntry(
	_In_opt_ LPMDB lpMDB,
	_In_opt_ LPADRBOOK lpAB,
	_In_opt_ LPMAPICONTAINER lpContainer,
	_In_opt_ LPMAPISESSION lpMAPISession,
	ULONG cbEntryID,
	_In_opt_ LPENTRYID lpEntryID,
	_In_opt_ LPCIID lpInterface,
	ULONG ulFlags,
	_Out_opt_ ULONG* ulObjTypeRet, // optional - can be NULL
	_Deref_out_opt_ LPUNKNOWN* lppUnk) // required
{
	if (!lppUnk) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;
	ULONG ulObjType = 0;
	LPUNKNOWN lpUnk = nullptr;
	ULONG ulNoCacheFlags = 0;

	*lppUnk = nullptr;

	if (RegKeys[regKeyMAPI_NO_CACHE].ulCurDWORD)
	{
		ulFlags |= MAPI_NO_CACHE;
	}

	// in case we need to retry without MAPI_NO_CACHE - do not add MAPI_NO_CACHE to ulFlags after this point
	if (MAPI_NO_CACHE & ulFlags) ulNoCacheFlags = ulFlags & ~MAPI_NO_CACHE;

	if (lpInterface && fIsSet(DBGGeneric))
	{
		auto szGuid = GUIDToStringAndName(lpInterface);
		DebugPrint(DBGGeneric, L"CallOpenEntry: OpenEntry asking for %ws\n", szGuid.c_str());
	}

	if (lpMDB)
	{
		DebugPrint(DBGGeneric, L"CallOpenEntry: Calling OpenEntry on MDB with ulFlags = 0x%X\n", ulFlags);
		WC_MAPI(lpMDB->OpenEntry(
			cbEntryID,
			lpEntryID,
			lpInterface,
			ulFlags,
			&ulObjType,
			&lpUnk));
		if (MAPI_E_UNKNOWN_FLAGS == hRes && ulNoCacheFlags)
		{
			DebugPrint(DBGGeneric, L"CallOpenEntry 2nd attempt: Calling OpenEntry on MDB with ulFlags = 0x%X\n", ulNoCacheFlags);
			hRes = S_OK;
			if (lpUnk) lpUnk->Release();
			lpUnk = nullptr;
			WC_MAPI(lpMDB->OpenEntry(
				cbEntryID,
				lpEntryID,
				lpInterface,
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
	if (lpAB && !lpUnk)
	{
		hRes = S_OK;
		DebugPrint(DBGGeneric, L"CallOpenEntry: Calling OpenEntry on AB with ulFlags = 0x%X\n", ulFlags);
		WC_MAPI(lpAB->OpenEntry(
			cbEntryID,
			lpEntryID,
			NULL, // no interface
			ulFlags,
			&ulObjType,
			&lpUnk));
		if (MAPI_E_UNKNOWN_FLAGS == hRes && ulNoCacheFlags)
		{
			DebugPrint(DBGGeneric, L"CallOpenEntry 2nd attempt: Calling OpenEntry on AB with ulFlags = 0x%X\n", ulNoCacheFlags);
			hRes = S_OK;
			if (lpUnk) lpUnk->Release();
			lpUnk = nullptr;
			WC_MAPI(lpAB->OpenEntry(
				cbEntryID,
				lpEntryID,
				NULL, // no interface
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
		hRes = S_OK;
		DebugPrint(DBGGeneric, L"CallOpenEntry: Calling OpenEntry on Container with ulFlags = 0x%X\n", ulFlags);
		WC_MAPI(lpContainer->OpenEntry(
			cbEntryID,
			lpEntryID,
			lpInterface,
			ulFlags,
			&ulObjType,
			&lpUnk));
		if (MAPI_E_UNKNOWN_FLAGS == hRes && ulNoCacheFlags)
		{
			DebugPrint(DBGGeneric, L"CallOpenEntry 2nd attempt: Calling OpenEntry on Container with ulFlags = 0x%X\n", ulNoCacheFlags);
			hRes = S_OK;
			if (lpUnk) lpUnk->Release();
			lpUnk = nullptr;
			WC_MAPI(lpContainer->OpenEntry(
				cbEntryID,
				lpEntryID,
				lpInterface,
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

	if (lpMAPISession && !lpUnk)
	{
		hRes = S_OK;
		DebugPrint(DBGGeneric, L"CallOpenEntry: Calling OpenEntry on Session with ulFlags = 0x%X\n", ulFlags);
		WC_MAPI(lpMAPISession->OpenEntry(
			cbEntryID,
			lpEntryID,
			lpInterface,
			ulFlags,
			&ulObjType,
			&lpUnk));
		if (MAPI_E_UNKNOWN_FLAGS == hRes && ulNoCacheFlags)
		{
			DebugPrint(DBGGeneric, L"CallOpenEntry 2nd attempt: Calling OpenEntry on Session with ulFlags = 0x%X\n", ulNoCacheFlags);
			hRes = S_OK;
			if (lpUnk) lpUnk->Release();
			lpUnk = nullptr;
			WC_MAPI(lpMAPISession->OpenEntry(
				cbEntryID,
				lpEntryID,
				lpInterface,
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

	if (lpUnk)
	{
		auto szFlags = InterpretNumberAsStringProp(ulObjType, PR_OBJECT_TYPE);
		DebugPrint(DBGGeneric, L"OnOpenEntryID: Got object of type 0x%08X = %ws\n", ulObjType, szFlags.c_str());
		*lppUnk = lpUnk;
	}
	if (ulObjTypeRet) *ulObjTypeRet = ulObjType;
	return hRes;
}

_Check_return_ HRESULT CallOpenEntry(
	_In_opt_ LPMDB lpMDB,
	_In_opt_ LPADRBOOK lpAB,
	_In_opt_ LPMAPICONTAINER lpContainer,
	_In_opt_ LPMAPISESSION lpMAPISession,
	_In_opt_ LPSBinary lpSBinary,
	_In_opt_ LPCIID lpInterface,
	ULONG ulFlags,
	_Out_opt_ ULONG* ulObjTypeRet,
	_Deref_out_opt_ LPUNKNOWN* lppUnk)
{
	auto hRes = S_OK;
	WC_H(CallOpenEntry(
		lpMDB,
		lpAB,
		lpContainer,
		lpMAPISession,
		lpSBinary ? lpSBinary->cb : 0,
		reinterpret_cast<LPENTRYID>(lpSBinary ? lpSBinary->lpb : 0),
		lpInterface,
		ulFlags,
		ulObjTypeRet,
		lppUnk));
	return hRes;
}

// Concatenate two property arrays without duplicates
_Check_return_ HRESULT ConcatSPropTagArrays(
	_In_ LPSPropTagArray lpArray1,
	_In_opt_ LPSPropTagArray lpArray2,
	_Deref_out_opt_ LPSPropTagArray *lpNewArray)
{
	if (!lpNewArray) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;
	LPSPropTagArray lpLocalArray = nullptr;

	*lpNewArray = nullptr;

	// Add the sizes of the passed in arrays (0 if they were NULL)
	auto iNewArraySize = lpArray1 ? lpArray1->cValues : 0;

	if (lpArray2 && lpArray1)
	{
		for (ULONG iSourceArray = 0; iSourceArray < lpArray2->cValues; iSourceArray++)
		{
			if (!IsDuplicateProp(lpArray1, lpArray2->aulPropTag[iSourceArray]))
			{
				iNewArraySize++;
			}
		}
	}
	else
	{
		iNewArraySize = iNewArraySize + (lpArray2 ? lpArray2->cValues : 0);
	}

	if (!iNewArraySize) return MAPI_E_CALL_FAILED;

	// Allocate memory for the new prop tag array
	EC_H(MAPIAllocateBuffer(
		CbNewSPropTagArray(iNewArraySize),
		reinterpret_cast<LPVOID*>(&lpLocalArray)));

	if (lpLocalArray)
	{
		ULONG iTargetArray = 0;
		if (lpArray1)
		{
			for (ULONG iSourceArray = 0; iSourceArray < lpArray1->cValues; iSourceArray++)
			{
				if (PROP_TYPE(lpArray1->aulPropTag[iSourceArray]) != PT_NULL) // ditch bad props
				{
					lpLocalArray->aulPropTag[iTargetArray++] = lpArray1->aulPropTag[iSourceArray];
				}
			}
		}

		if (lpArray2)
		{
			for (ULONG iSourceArray = 0; iSourceArray < lpArray2->cValues; iSourceArray++)
			{
				if (PROP_TYPE(lpArray2->aulPropTag[iSourceArray]) != PT_NULL) // ditch bad props
				{
					if (!IsDuplicateProp(lpArray1, lpArray2->aulPropTag[iSourceArray]))
					{
						lpLocalArray->aulPropTag[iTargetArray++] = lpArray2->aulPropTag[iSourceArray];
					}
				}
			}
		}

		// <= since we may have thrown some PT_NULL tags out - just make sure we didn't overrun.
		EC_H(iTargetArray <= iNewArraySize ? S_OK : MAPI_E_CALL_FAILED);

		// since we may have ditched some tags along the way, reset our size
		lpLocalArray->cValues = iTargetArray;

		if (FAILED(hRes))
		{
			MAPIFreeBuffer(lpLocalArray);
		}
		else
		{
			*lpNewArray = static_cast<LPSPropTagArray>(lpLocalArray);
		}
	}

	return hRes;
}

// May not behave correctly if lpSrcFolder == lpDestFolder
// We can check that the pointers aren't equal, but they could be different
// and still refer to the same folder.
_Check_return_ HRESULT CopyFolderContents(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bCopyAssociatedContents, bool bMove, bool bSingleCall, _In_ HWND hWnd)
{
	auto hRes = S_OK;
	LPMAPITABLE lpSrcContents = nullptr;
	LPSRowSet pRows = nullptr;

	enum
	{
		fldPR_ENTRYID,
		fldNUM_COLS
	};

	static const SizedSPropTagArray(fldNUM_COLS, fldCols) =
	{
	fldNUM_COLS,
	PR_ENTRYID,
	};

	DebugPrint(DBGGeneric, L"CopyFolderContents: lpSrcFolder = %p, lpDestFolder = %p, bCopyAssociatedContents = %d, bMove = %d\n",
		lpSrcFolder,
		lpDestFolder,
		bCopyAssociatedContents,
		bMove
	);

	if (!lpSrcFolder || !lpDestFolder) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpSrcFolder->GetContentsTable(
		fMapiUnicode | (bCopyAssociatedContents ? MAPI_ASSOCIATED : NULL),
		&lpSrcContents));

	if (lpSrcContents)
	{
		EC_MAPI(lpSrcContents->SetColumns(LPSPropTagArray(&fldCols), TBL_BATCH));

		ULONG ulRowCount = 0;
		EC_MAPI(lpSrcContents->GetRowCount(0, &ulRowCount));

		if (bSingleCall && ulRowCount < ULONG_MAX / sizeof(SBinary))
		{
			SBinaryArray sbaEID = { 0 };
			sbaEID.cValues = ulRowCount;
			EC_H(MAPIAllocateBuffer(sizeof(SBinary)* ulRowCount, reinterpret_cast<LPVOID*>(&sbaEID.lpbin)));
			ZeroMemory(sbaEID.lpbin, sizeof(SBinary)* ulRowCount);

			if (!FAILED(hRes)) for (ULONG ulRowsCopied = 0; ulRowsCopied < ulRowCount; ulRowsCopied++)
			{
				hRes = S_OK;
				if (pRows) FreeProws(pRows);
				pRows = nullptr;
				EC_MAPI(lpSrcContents->QueryRows(
					1,
					NULL,
					&pRows));
				if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

				if (pRows && PT_ERROR != PROP_TYPE(pRows->aRow->lpProps[fldPR_ENTRYID].ulPropTag))
				{
					EC_H(CopySBinary(&sbaEID.lpbin[ulRowsCopied], &pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin, sbaEID.lpbin));
				}
			}

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK

			auto ulCopyFlags = bMove ? MESSAGE_MOVE : 0;

			if (lpProgress)
				ulCopyFlags |= MESSAGE_DIALOG;

			EC_MAPI(lpSrcFolder->CopyMessages(
				&sbaEID,
				&IID_IMAPIFolder,
				lpDestFolder,
				lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
				lpProgress,
				ulCopyFlags));
			MAPIFreeBuffer(sbaEID.lpbin);

			if (lpProgress) lpProgress->Release();
		}
		else
		{
			if (!FAILED(hRes)) for (ULONG ulRowsCopied = 0; ulRowsCopied < ulRowCount; ulRowsCopied++)
			{
				hRes = S_OK;
				if (pRows) FreeProws(pRows);
				pRows = nullptr;
				EC_MAPI(lpSrcContents->QueryRows(
					1,
					NULL,
					&pRows));
				if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

				if (pRows && PT_ERROR != PROP_TYPE(pRows->aRow->lpProps[fldPR_ENTRYID].ulPropTag))
				{
					SBinaryArray sbaEID = { 0 };
					DebugPrint(DBGGeneric, L"Source Message =\n");
					DebugPrintBinary(DBGGeneric, &pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin);

					sbaEID.cValues = 1;
					sbaEID.lpbin = &pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin;

					LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK

					auto ulCopyFlags = bMove ? MESSAGE_MOVE : 0;

					if (lpProgress)
						ulCopyFlags |= MESSAGE_DIALOG;

					EC_MAPI(lpSrcFolder->CopyMessages(
						&sbaEID,
						&IID_IMAPIFolder,
						lpDestFolder,
						lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
						lpProgress,
						ulCopyFlags));

					if (S_OK == hRes) DebugPrint(DBGGeneric, L"Message Copied\n");

					if (lpProgress) lpProgress->Release();
				}

				if (S_OK != hRes) DebugPrint(DBGGeneric, L"Message Copy Failed\n");
			}
		}

		lpSrcContents->Release();
	}

	if (pRows) FreeProws(pRows);
	return hRes;
}

_Check_return_ HRESULT CopyFolderRules(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bReplace)
{
	if (!lpSrcFolder || !lpDestFolder) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;
	LPEXCHANGEMODIFYTABLE lpSrcTbl = nullptr;
	LPEXCHANGEMODIFYTABLE lpDstTbl = nullptr;

	EC_MAPI(lpSrcFolder->OpenProperty(
		PR_RULES_TABLE,
		const_cast<LPGUID>(&IID_IExchangeModifyTable),
		0,
		MAPI_DEFERRED_ERRORS,
		reinterpret_cast<LPUNKNOWN*>(&lpSrcTbl)));

	EC_MAPI(lpDestFolder->OpenProperty(
		PR_RULES_TABLE,
		const_cast<LPGUID>(&IID_IExchangeModifyTable),
		0,
		MAPI_DEFERRED_ERRORS,
		reinterpret_cast<LPUNKNOWN*>(&lpDstTbl)));

	if (lpSrcTbl && lpDstTbl)
	{
		LPMAPITABLE lpTable = nullptr;
		lpSrcTbl->GetTable(0, &lpTable);

		if (lpTable)
		{
			static const SizedSPropTagArray(9, ruleTags) =
			{
			9,
			PR_RULE_ACTIONS,
			PR_RULE_CONDITION,
			PR_RULE_LEVEL,
			PR_RULE_NAME,
			PR_RULE_PROVIDER,
			PR_RULE_PROVIDER_DATA,
			PR_RULE_SEQUENCE,
			PR_RULE_STATE,
			PR_RULE_USER_FLAGS,
			};

			EC_MAPI(lpTable->SetColumns(LPSPropTagArray(&ruleTags), 0));

			LPSRowSet lpRows = nullptr;

			EC_MAPI(HrQueryAllRows(
				lpTable,
				NULL,
				NULL,
				NULL,
				NULL,
				&lpRows));

			if (lpRows && lpRows->cRows < MAXNewROWLIST)
			{
				LPROWLIST lpTempList = nullptr;

				EC_H(MAPIAllocateBuffer(CbNewROWLIST(lpRows->cRows), reinterpret_cast<LPVOID*>(&lpTempList)));

				if (lpTempList)
				{
					lpTempList->cEntries = lpRows->cRows;

					for (ULONG iArrayPos = 0; iArrayPos < lpRows->cRows; iArrayPos++)
					{
						lpTempList->aEntries[iArrayPos].ulRowFlags = ROW_ADD;
						EC_H(MAPIAllocateMore(
							lpRows->aRow[iArrayPos].cValues * sizeof(SPropValue),
							lpTempList,
							reinterpret_cast<LPVOID*>(&lpTempList->aEntries[iArrayPos].rgPropVals)));
						if (SUCCEEDED(hRes) && lpTempList->aEntries[iArrayPos].rgPropVals)
						{
							ULONG ulDst = 0;
							for (ULONG ulSrc = 0; ulSrc < lpRows->aRow[iArrayPos].cValues; ulSrc++)
							{
								if (lpRows->aRow[iArrayPos].lpProps[ulSrc].ulPropTag == PR_RULE_PROVIDER_DATA)
								{
									if (!lpRows->aRow[iArrayPos].lpProps[ulSrc].Value.bin.cb ||
										!lpRows->aRow[iArrayPos].lpProps[ulSrc].Value.bin.lpb)
									{
										// PR_RULE_PROVIDER_DATA was NULL - we don't want this
										continue;
									}
								}

								EC_H(MyPropCopyMore(
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

					EC_MAPI(lpDstTbl->ModifyTable(ulFlags, lpTempList));

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
_Check_return_ HRESULT CopyPropertyAsStream(_In_ LPMAPIPROP lpSourcePropObj,
	_In_ LPMAPIPROP lpTargetPropObj,
	ULONG ulSourceTag,
	ULONG ulTargetTag)
{
	auto hRes = S_OK;
	LPSTREAM lpStmSource = nullptr;
	LPSTREAM lpStmTarget = nullptr;
	LARGE_INTEGER li;
	ULARGE_INTEGER uli;
	ULARGE_INTEGER ulBytesRead;
	ULARGE_INTEGER ulBytesWritten;

	if (!lpSourcePropObj || !lpTargetPropObj || !ulSourceTag || !ulTargetTag) return MAPI_E_INVALID_PARAMETER;
	if (PROP_TYPE(ulSourceTag) != PROP_TYPE(ulTargetTag)) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpSourcePropObj->OpenProperty(
		ulSourceTag,
		&IID_IStream,
		STGM_READ | STGM_DIRECT,
		NULL,
		reinterpret_cast<LPUNKNOWN *>(&lpStmSource)));

	EC_MAPI(lpTargetPropObj->OpenProperty(
		ulTargetTag,
		&IID_IStream,
		STGM_READWRITE | STGM_DIRECT,
		MAPI_CREATE | MAPI_MODIFY,
		reinterpret_cast<LPUNKNOWN *>(&lpStmTarget)));

	if (lpStmSource && lpStmTarget)
	{
		li.QuadPart = 0;
		uli.QuadPart = MAXLONGLONG;

		EC_MAPI(lpStmSource->Seek(li, STREAM_SEEK_SET, NULL));

		EC_MAPI(lpStmTarget->Seek(li, STREAM_SEEK_SET, NULL));

		EC_MAPI(lpStmSource->CopyTo(lpStmTarget, uli, &ulBytesRead, &ulBytesWritten));

		// This may not be necessary since we opened with STGM_DIRECT
		EC_MAPI(lpStmTarget->Commit(STGC_DEFAULT));
	}

	if (lpStmTarget) lpStmTarget->Release();
	if (lpStmSource) lpStmSource->Release();
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
// CopySBinary()
//
// Parameters
// psbDest - Address of the destination binary
// psbSrc - Address of the source binary
// lpParent - Pointer to parent object (not, however, pointer to pointer!)
//
// Purpose
// Allocates a new SBinary and copies psbSrc into it
_Check_return_ HRESULT CopySBinary(_Out_ LPSBinary psbDest, _In_ const LPSBinary psbSrc, _In_ LPVOID lpParent)
{
	auto hRes = S_OK;

	if (!psbDest || !psbSrc) return MAPI_E_INVALID_PARAMETER;

	psbDest->cb = psbSrc->cb;

	if (psbSrc->cb)
	{
		if (lpParent)
		{
			EC_H(MAPIAllocateMore(
				psbSrc->cb,
				lpParent,
				reinterpret_cast<LPVOID *>(&psbDest->lpb)))
		}
		else
		{
			EC_H(MAPIAllocateBuffer(
				psbSrc->cb,
				reinterpret_cast<LPVOID *>(&psbDest->lpb)));
		}

		if (S_OK == hRes)
			CopyMemory(psbDest->lpb, psbSrc->lpb, psbSrc->cb);
	}

	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
// CopyString()
//
// Parameters
// lpszDestination - Address of pointer to destination string
// szSource - Pointer to source string
// lpParent - Pointer to parent object (not, however, pointer to pointer!)
//
// Purpose
// Uses MAPI to allocate a new string (szDestination) and copy szSource into it
// Uses lpParent as the parent for MAPIAllocateMore if possible
_Check_return_ HRESULT CopyStringA(_Deref_out_z_ LPSTR* lpszDestination, _In_z_ LPCSTR szSource, _In_opt_ LPVOID pParent)
{
	auto hRes = S_OK;
	size_t cbSource = 0;

	if (!szSource)
	{
		*lpszDestination = nullptr;
		return hRes;
	}

	EC_H(StringCbLengthA(szSource, STRSAFE_MAX_CCH * sizeof(char), &cbSource));
	cbSource += sizeof(char);

	if (pParent)
	{
		EC_H(MAPIAllocateMore(
			static_cast<ULONG>(cbSource),
			pParent,
			reinterpret_cast<LPVOID*>(lpszDestination)));
	}
	else
	{
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(cbSource),
			reinterpret_cast<LPVOID*>(lpszDestination)));
	}

	EC_H(StringCbCopyA(*lpszDestination, cbSource, szSource));

	return hRes;
}

_Check_return_ HRESULT CopyStringW(_Deref_out_z_ LPWSTR* lpszDestination, _In_z_ LPCWSTR szSource, _In_opt_ LPVOID pParent)
{
	auto hRes = S_OK;
	size_t cbSource = 0;

	if (!szSource)
	{
		*lpszDestination = nullptr;
		return hRes;
	}

	EC_H(StringCbLengthW(szSource, STRSAFE_MAX_CCH * sizeof(WCHAR), &cbSource));
	cbSource += sizeof(WCHAR);

	if (pParent)
	{
		EC_H(MAPIAllocateMore(
			static_cast<ULONG>(cbSource),
			pParent,
			reinterpret_cast<LPVOID*>(lpszDestination)));
	}
	else
	{
		EC_H(MAPIAllocateBuffer(
			static_cast<ULONG>(cbSource),
			reinterpret_cast<LPVOID*>(lpszDestination)));
	}

	EC_H(StringCbCopyW(*lpszDestination, cbSource, szSource));

	return hRes;
}

// Allocates and creates a restriction that looks for existence of
// a particular property that matches the given string
// If lpParent is passed in, it is used as the allocation parent.
_Check_return_ HRESULT CreatePropertyStringRestriction(ULONG ulPropTag,
	_In_ wstring& szString,
	ULONG ulFuzzyLevel,
	_In_opt_ LPVOID lpParent,
	_Deref_out_opt_ LPSRestriction* lppRes)
{
	if (PROP_TYPE(ulPropTag) != PT_UNICODE) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;
	LPSRestriction lpRes = nullptr;
	LPSRestriction lpResLevel1 = nullptr;
	LPSPropValue lpspvSubject = nullptr;
	LPVOID lpAllocationParent = nullptr;

	*lppRes = nullptr;

	if (szString.empty()) return MAPI_E_INVALID_PARAMETER;

	// Allocate and create our SRestriction
	// Allocate base memory:
	if (lpParent)
	{
		EC_H(MAPIAllocateMore(
			sizeof(SRestriction),
			lpParent,
			reinterpret_cast<LPVOID*>(&lpRes)));

		lpAllocationParent = lpParent;
	}
	else
	{
		EC_H(MAPIAllocateBuffer(
			sizeof(SRestriction),
			reinterpret_cast<LPVOID*>(&lpRes)));

		lpAllocationParent = lpRes;
	}

	EC_H(MAPIAllocateMore(
		sizeof(SRestriction) * 2,
		lpAllocationParent,
		reinterpret_cast<LPVOID*>(&lpResLevel1)));

	EC_H(MAPIAllocateMore(
		sizeof(SPropValue),
		lpAllocationParent,
		reinterpret_cast<LPVOID*>(&lpspvSubject)));

	if (lpRes && lpResLevel1 && lpspvSubject)
	{
		// Zero out allocated memory.
		ZeroMemory(lpRes, sizeof(SRestriction));
		ZeroMemory(lpResLevel1, sizeof(SRestriction) * 2);
		ZeroMemory(lpspvSubject, sizeof(SPropValue));

		// Root Node
		lpRes->rt = RES_AND;
		lpRes->res.resAnd.cRes = 2;
		lpRes->res.resAnd.lpRes = lpResLevel1;

		lpResLevel1[0].rt = RES_EXIST;
		lpResLevel1[0].res.resExist.ulPropTag = ulPropTag;
		lpResLevel1[0].res.resExist.ulReserved1 = 0;
		lpResLevel1[0].res.resExist.ulReserved2 = 0;

		lpResLevel1[1].rt = RES_CONTENT;
		lpResLevel1[1].res.resContent.ulPropTag = ulPropTag;
		lpResLevel1[1].res.resContent.ulFuzzyLevel = ulFuzzyLevel;
		lpResLevel1[1].res.resContent.lpProp = lpspvSubject;

		// Allocate and fill out properties:
		lpspvSubject->ulPropTag = ulPropTag;

		EC_H(CopyStringW(
			&lpspvSubject->Value.lpszW,
			szString.c_str(),
			lpAllocationParent));

		DebugPrint(DBGGeneric, L"CreatePropertyStringRestriction built restriction:\n");
		DebugPrintRestriction(DBGGeneric, lpRes, NULL);

		*lppRes = lpRes;
	}

	if (hRes != S_OK)
	{
		DebugPrint(DBGGeneric, L"Failed to create restriction\n");
		MAPIFreeBuffer(lpRes);
		*lppRes = nullptr;
	}

	return hRes;
}

_Check_return_ HRESULT CreateRangeRestriction(ULONG ulPropTag,
	_In_ wstring& szString,
	_In_opt_ LPVOID lpParent,
	_Deref_out_opt_ LPSRestriction* lppRes)
{
	auto hRes = S_OK;
	LPSRestriction lpRes = nullptr;
	LPSRestriction lpResLevel1 = nullptr;
	LPSPropValue lpspvSubject = nullptr;
	LPVOID lpAllocationParent = nullptr;

	*lppRes = nullptr;

	if (szString.empty()) return MAPI_E_INVALID_PARAMETER;
	if (PROP_TYPE(ulPropTag) != PT_UNICODE) return MAPI_E_INVALID_PARAMETER;

	// Allocate and create our SRestriction
	// Allocate base memory:
	if (lpParent)
	{
		EC_H(MAPIAllocateMore(
			sizeof(SRestriction),
			lpParent,
			reinterpret_cast<LPVOID*>(&lpRes)));

		lpAllocationParent = lpParent;
	}
	else
	{
		EC_H(MAPIAllocateBuffer(
			sizeof(SRestriction),
			reinterpret_cast<LPVOID*>(&lpRes)));

		lpAllocationParent = lpRes;
	}

	EC_H(MAPIAllocateMore(
		sizeof(SRestriction) * 2,
		lpAllocationParent,
		reinterpret_cast<LPVOID*>(&lpResLevel1)));

	EC_H(MAPIAllocateMore(
		sizeof(SPropValue),
		lpAllocationParent,
		reinterpret_cast<LPVOID*>(&lpspvSubject)));

	if (lpRes && lpResLevel1 && lpspvSubject)
	{
		// Zero out allocated memory.
		ZeroMemory(lpRes, sizeof(SRestriction));
		ZeroMemory(lpResLevel1, sizeof(SRestriction) * 2);
		ZeroMemory(lpspvSubject, sizeof(SPropValue));

		// Root Node
		lpRes->rt = RES_AND;
		lpRes->res.resAnd.cRes = 2;
		lpRes->res.resAnd.lpRes = lpResLevel1;

		lpResLevel1[0].rt = RES_EXIST;
		lpResLevel1[0].res.resExist.ulPropTag = ulPropTag;
		lpResLevel1[0].res.resExist.ulReserved1 = 0;
		lpResLevel1[0].res.resExist.ulReserved2 = 0;

		lpResLevel1[1].rt = RES_PROPERTY;
		lpResLevel1[1].res.resProperty.ulPropTag = ulPropTag;
		lpResLevel1[1].res.resProperty.relop = RELOP_GE;
		lpResLevel1[1].res.resProperty.lpProp = lpspvSubject;

		// Allocate and fill out properties:
		lpspvSubject->ulPropTag = ulPropTag;

		EC_H(CopyStringW(
			&lpspvSubject->Value.lpszW,
			szString.c_str(),
			lpAllocationParent));

		DebugPrint(DBGGeneric, L"CreateRangeRestriction built restriction:\n");
		DebugPrintRestriction(DBGGeneric, lpRes, NULL);

		*lppRes = lpRes;
	}

	if (hRes != S_OK)
	{
		DebugPrint(DBGGeneric, L"Failed to create restriction\n");
		MAPIFreeBuffer(lpRes);
		*lppRes = nullptr;
	}

	return hRes;
}

_Check_return_ HRESULT DeleteProperty(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
{
	auto hRes = S_OK;
	SPropTagArray ptaTag;
	LPSPropProblemArray pProbArray = nullptr;

	if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

	if (PT_ERROR == PROP_TYPE(ulPropTag))
		ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED);

	ptaTag.cValues = 1;
	ptaTag.aulPropTag[0] = ulPropTag;

	DebugPrint(DBGGeneric, L"DeleteProperty: Deleting prop 0x%08X from MAPI item %p.\n", ulPropTag, lpMAPIProp);

	EC_MAPI(lpMAPIProp->DeleteProps(
		&ptaTag,
		&pProbArray));

	if (S_OK == hRes && pProbArray)
	{
		EC_PROBLEMARRAY(pProbArray);
	}

	if (SUCCEEDED(hRes))
	{
		EC_MAPI(lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
	}

	MAPIFreeBuffer(pProbArray);

	return hRes;
}

// Delete items to the wastebasket of the passed in mdb, if it exists.
_Check_return_ HRESULT DeleteToDeletedItems(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpSourceFolder, _In_ LPENTRYLIST lpEIDs, _In_ HWND hWnd)
{
	auto hRes = S_OK;
	LPMAPIFOLDER lpWasteFolder = nullptr;

	if (!lpMDB || !lpSourceFolder || !lpEIDs) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"DeleteToDeletedItems: Deleting from folder %p in store %p\n",
		lpSourceFolder,
		lpMDB);

	WC_H(OpenDefaultFolder(DEFAULT_DELETEDITEMS, lpMDB, &lpWasteFolder));

	if (lpWasteFolder)
	{
		LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK

		auto ulCopyFlags = MESSAGE_MOVE;

		if (lpProgress)
			ulCopyFlags |= MESSAGE_DIALOG;

		EC_MAPI(lpSourceFolder->CopyMessages(
			lpEIDs,
			NULL, // default interface
			lpWasteFolder,
			lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
			lpProgress,
			ulCopyFlags));

		if (lpProgress) lpProgress->Release();
	}

	if (lpWasteFolder) lpWasteFolder->Release();
	return hRes;
}

_Check_return_ bool FindPropInPropTagArray(_In_ LPSPropTagArray lpspTagArray, ULONG ulPropToFind, _Out_ ULONG* lpulRowFound)
{
	*lpulRowFound = 0;
	if (!lpspTagArray) return false;
	for (ULONG i = 0; i < lpspTagArray->cValues; i++)
	{
		if (PROP_ID(ulPropToFind) == PROP_ID(lpspTagArray->aulPropTag[i]))
		{
			*lpulRowFound = i;
			return true;
		}
	}

	return false;
}

// See list of types (like MAPI_FOLDER) in mapidefs.h
_Check_return_ ULONG GetMAPIObjectType(_In_opt_ LPMAPIPROP lpMAPIProp)
{
	auto hRes = S_OK;
	ULONG ulObjType = 0;
	LPSPropValue lpProp = nullptr;

	if (!lpMAPIProp) return 0; // 0's not a valid Object type

	WC_MAPI(HrGetOneProp(
		lpMAPIProp,
		PR_OBJECT_TYPE,
		&lpProp));

	if (lpProp)
		ulObjType = lpProp->Value.ul;

	MAPIFreeBuffer(lpProp);
	return ulObjType;
}

_Check_return_ HRESULT GetInbox(_In_ LPMDB lpMDB, _Out_opt_ ULONG* lpcbeid, _Deref_out_opt_ LPENTRYID* lppeid)
{
	auto hRes = S_OK;
	ULONG cbInboxEID = 0;
	LPENTRYID lpInboxEID = nullptr;

	DebugPrint(DBGGeneric, L"GetInbox: getting Inbox\n");

	if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpMDB->GetReceiveFolder(
		LPTSTR("IPM.Note"), // STRING_OK this is the class of message we want
		fMapiUnicode, // flags
		&cbInboxEID, // size and...
		&lpInboxEID, // value of entry ID
		NULL)); // returns a message class if not NULL

	if (cbInboxEID && lpInboxEID)
	{
		WC_H(MAPIAllocateBuffer(cbInboxEID, reinterpret_cast<LPVOID*>(lppeid)));
		if (SUCCEEDED(hRes))
		{
			*lpcbeid = cbInboxEID;
			CopyMemory(*lppeid, lpInboxEID, *lpcbeid);
		}
	}

	MAPIFreeBuffer(lpInboxEID);
	return hRes;
}

_Check_return_ HRESULT GetInbox(_In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER* lpInbox)
{
	auto hRes = S_OK;
	ULONG cbInboxEID = 0;
	LPENTRYID lpInboxEID = nullptr;

	DebugPrint(DBGGeneric, L"GetInbox: getting Inbox from %p\n", lpMDB);

	*lpInbox = nullptr;

	if (!lpMDB || !lpInbox) return MAPI_E_INVALID_PARAMETER;

	EC_H(GetInbox(lpMDB, &cbInboxEID, &lpInboxEID));

	if (cbInboxEID && lpInboxEID)
	{
		// Get the Inbox...
		WC_H(CallOpenEntry(
			lpMDB,
			NULL,
			NULL,
			NULL,
			cbInboxEID,
			lpInboxEID,
			NULL,
			MAPI_BEST_ACCESS,
			NULL,
			reinterpret_cast<LPUNKNOWN*>(lpInbox)));
	}

	MAPIFreeBuffer(lpInboxEID);
	return hRes;
}

_Check_return_ HRESULT GetParentFolder(_In_ LPMAPIFOLDER lpChildFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER* lpParentFolder)
{
	auto hRes = S_OK;
	ULONG cProps;
	LPSPropValue lpProps = nullptr;

	*lpParentFolder = nullptr;

	if (!lpChildFolder) return MAPI_E_INVALID_PARAMETER;

	enum
	{
		PARENTEID,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptaSrcFolder) =
	{
	NUM_COLS,
	PR_PARENT_ENTRYID
	};

	// Get PR_PARENT_ENTRYID
	EC_H_GETPROPS(lpChildFolder->GetProps(
		LPSPropTagArray(&sptaSrcFolder),
		fMapiUnicode,
		&cProps,
		&lpProps));

	if (lpProps && PT_ERROR != PROP_TYPE(lpProps[PARENTEID].ulPropTag))
	{
		WC_H(CallOpenEntry(
			lpMDB,
			NULL,
			NULL,
			NULL,
			lpProps[PARENTEID].Value.bin.cb,
			reinterpret_cast<LPENTRYID>(lpProps[PARENTEID].Value.bin.lpb),
			NULL,
			MAPI_BEST_ACCESS,
			NULL,
			reinterpret_cast<LPUNKNOWN*>(lpParentFolder)));
	}

	MAPIFreeBuffer(lpProps);
	return hRes;
}

_Check_return_ HRESULT GetPropsNULL(_In_ LPMAPIPROP lpMAPIProp, ULONG ulFlags, _Out_ ULONG* lpcValues, _Deref_out_opt_ LPSPropValue* lppPropArray)
{
	auto hRes = S_OK;
	*lpcValues = NULL;
	*lppPropArray = nullptr;

	if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;
	LPSPropTagArray lpTags = nullptr;
	if (RegKeys[regkeyUSE_GETPROPLIST].ulCurDWORD)
	{
		DebugPrint(DBGGeneric, L"GetPropsNULL: Calling GetPropList\n");
		WC_MAPI(lpMAPIProp->GetPropList(
			ulFlags,
			&lpTags));

		if (MAPI_E_BAD_CHARWIDTH == hRes)
		{
			hRes = S_OK;
			EC_MAPI(lpMAPIProp->GetPropList(
				NULL,
				&lpTags));
		}
		else
		{
			CHECKHRESMSG(hRes, IDS_NOPROPLIST);
		}
	}
	else
	{
		DebugPrint(DBGGeneric, L"GetPropsNULL: Calling GetProps(NULL) on %p\n", lpMAPIProp);
	}

	WC_H_GETPROPS(lpMAPIProp->GetProps(
		lpTags,
		ulFlags,
		lpcValues,
		lppPropArray));
	MAPIFreeBuffer(lpTags);

	return hRes;
}

_Check_return_ HRESULT GetSpecialFolderEID(_In_ LPMDB lpMDB, ULONG ulFolderPropTag, _Out_opt_ ULONG* lpcbeid, _Deref_out_opt_ LPENTRYID* lppeid)
{
	auto hRes = S_OK;
	LPMAPIFOLDER lpInbox = nullptr;

	DebugPrint(DBGGeneric, L"GetSpecialFolderEID: getting 0x%X from %p\n", ulFolderPropTag, lpMDB);

	if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;

	LPSPropValue lpProp = nullptr;
	WC_H(GetInbox(lpMDB, &lpInbox));
	if (lpInbox)
	{
		WC_H_MSG(HrGetOneProp(lpInbox, ulFolderPropTag, &lpProp),
			IDS_GETSPECIALFOLDERINBOXMISSINGPROP);
		lpInbox->Release();
	}

	if (!lpProp)
	{
		hRes = S_OK;
		LPMAPIFOLDER lpRootFolder = nullptr;
		// Open root container.
		EC_H(CallOpenEntry(
			lpMDB,
			NULL,
			NULL,
			NULL,
			NULL, // open root container
			NULL,
			MAPI_BEST_ACCESS,
			NULL,
			reinterpret_cast<LPUNKNOWN*>(&lpRootFolder)));
		if (lpRootFolder)
		{
			EC_H_MSG(HrGetOneProp(lpRootFolder, ulFolderPropTag, &lpProp),
				IDS_GETSPECIALFOLDERROOTMISSINGPROP);
			lpRootFolder->Release();
		}
	}

	if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag) && lpProp->Value.bin.cb)
	{
		WC_H(MAPIAllocateBuffer(lpProp->Value.bin.cb, reinterpret_cast<LPVOID*>(lppeid)));
		if (SUCCEEDED(hRes))
		{
			*lpcbeid = lpProp->Value.bin.cb;
			CopyMemory(*lppeid, lpProp->Value.bin.lpb, *lpcbeid);
		}
	}

	if (MAPI_E_NOT_FOUND == hRes)
	{
		DebugPrint(DBGGeneric, L"Special folder not found.\n");
	}

	MAPIFreeBuffer(lpProp);
	return hRes;
}

_Check_return_ HRESULT IsAttachmentBlocked(_In_ LPMAPISESSION lpMAPISession, _In_z_ LPCWSTR pwszFileName, _Out_ bool* pfBlocked)
{
	if (!lpMAPISession || !pwszFileName || !pfBlocked) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;
	IAttachmentSecurity* lpAttachSec = nullptr;
	BOOL bBlocked = false;

	EC_MAPI(lpMAPISession->QueryInterface(IID_IAttachmentSecurity, reinterpret_cast<LPVOID*>(&lpAttachSec)));
	if (SUCCEEDED(hRes) && lpAttachSec)
	{
		EC_MAPI(lpAttachSec->IsAttachmentBlocked(pwszFileName, &bBlocked));
	}

	if (lpAttachSec) lpAttachSec->Release();

	*pfBlocked = !!bBlocked;
	return hRes;
}

_Check_return_ bool IsDuplicateProp(_In_ LPSPropTagArray lpArray, ULONG ulPropTag)
{
	if (!lpArray) return false;

	for (ULONG i = 0; i < lpArray->cValues; i++)
	{
		// They're dupes if the IDs are the same
		if (RegKeys[regkeyALLOW_DUPE_COLUMNS].ulCurDWORD)
		{
			if (lpArray->aulPropTag[i] == ulPropTag)
				return true;
		}
		else
		{
			if (PROP_ID(lpArray->aulPropTag[i]) == PROP_ID(ulPropTag))
				return true;
		}
	}

	return false;
}

_Check_return_ HRESULT ManuallyEmptyFolder(_In_ LPMAPIFOLDER lpFolder, BOOL bAssoc, BOOL bHardDelete)
{
	if (!lpFolder) return MAPI_E_INVALID_PARAMETER;

	LPSRowSet pRows = nullptr;
	auto hRes = S_OK;
	ULONG iItemCount = 0;
	LPMAPITABLE lpContentsTable = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	enum
	{
		eidPR_ENTRYID,
		eidNUM_COLS
	};

	static const SizedSPropTagArray(eidNUM_COLS, eidCols) =
	{
	eidNUM_COLS,
	PR_ENTRYID,
	};

	// Get the table of contents of the folder
	WC_MAPI(lpFolder->GetContentsTable(
		bAssoc ? MAPI_ASSOCIATED : NULL,
		&lpContentsTable));

	if (SUCCEEDED(hRes) && lpContentsTable)
	{
		EC_MAPI(lpContentsTable->SetColumns(
			LPSPropTagArray(&eidCols),
			TBL_BATCH));

		// go to the first row
		EC_MAPI(lpContentsTable->SeekRow(
			BOOKMARK_BEGINNING,
			0,
			NULL));
		hRes = S_OK; // don't let failure here fail the whole op

		// get rows and delete messages one at a time (slow, but might work when batch deletion fails)
		if (!FAILED(hRes)) for (;;)
		{
			hRes = S_OK;
			if (pRows) FreeProws(pRows);
			pRows = nullptr;
			// Pull back a sizable block of rows to delete
			EC_MAPI(lpContentsTable->QueryRows(
				200,
				NULL,
				&pRows));
			if (FAILED(hRes) || !pRows || !pRows->cRows) break;

			for (ULONG iCurPropRow = 0; iCurPropRow < pRows->cRows; iCurPropRow++)
			{
				if (pRows->aRow[iCurPropRow].lpProps &&
					PR_ENTRYID == pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID].ulPropTag)
				{
					hRes = S_OK;
					ENTRYLIST eid = { 0 };
					eid.cValues = 1;
					eid.lpbin = &pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID].Value.bin;
					WC_MAPI(lpFolder->DeleteMessages(
						&eid,
						NULL,
						NULL,
						bHardDelete ? DELETE_HARD_DELETE : NULL));
					if (SUCCEEDED(hRes)) iItemCount++;
				}
			}
		}

		DebugPrint(DBGGeneric, L"ManuallyEmptyFolder deleted %u items\n", iItemCount);
	}

	if (pRows) FreeProws(pRows);
	if (lpContentsTable) lpContentsTable->Release();
	return hRes;
}

// Converts byte vector to LPBYTE allocated with MAPIAllocateMore
_Check_return_ LPBYTE ByteVectorToMAPI(vector<BYTE>& bin, LPVOID lpParent)
{
	if (bin.empty()) return nullptr;

	auto hRes = S_OK;
	LPBYTE lpBin = nullptr;

	// We allocate a couple extra bytes (initialized to NULL) in case this buffer is printed.
	WC_H(MAPIAllocateMore(
		static_cast<ULONG>(bin.size()) + sizeof(WCHAR),
		lpParent,
		reinterpret_cast<LPVOID*>(&lpBin)));
	if (SUCCEEDED(hRes) && lpBin)
	{
		memset(lpBin, 0, bin.size() + sizeof(WCHAR));
		memcpy(lpBin, &bin[0], bin.size());
		return lpBin;
	}

	return nullptr;
}

ULONG aulOneOffIDs[] = { dispidFormStorage,
dispidPageDirStream,
dispidFormPropStream,
dispidScriptStream,
dispidPropDefStream, // dispidPropDefStream must remain next to last in list
dispidCustomFlag }; // dispidCustomFlag must remain last in list

#define ulNumOneOffIDs (_countof(aulOneOffIDs))

_Check_return_ HRESULT RemoveOneOff(_In_ LPMESSAGE lpMessage, bool bRemovePropDef)
{
	if (!lpMessage) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGNamedProp, L"RemoveOneOff - removing one off named properties.\n");

	auto hRes = S_OK;
	MAPINAMEID rgnmid[ulNumOneOffIDs];
	LPMAPINAMEID rgpnmid[ulNumOneOffIDs];
	LPSPropTagArray lpTags = nullptr;

	for (ULONG i = 0; i < ulNumOneOffIDs; i++)
	{
		rgnmid[i].lpguid = const_cast<LPGUID>(&PSETID_Common);
		rgnmid[i].ulKind = MNID_ID;
		rgnmid[i].Kind.lID = aulOneOffIDs[i];
		rgpnmid[i] = &rgnmid[i];
	}

	EC_H(GetIDsFromNames(lpMessage,
		ulNumOneOffIDs,
		rgpnmid,
		0,
		&lpTags));
	if (lpTags)
	{
		LPSPropProblemArray lpProbArray = nullptr;

		DebugPrint(DBGNamedProp, L"RemoveOneOff - identified the following properties.\n");
		DebugPrintPropTagArray(DBGNamedProp, lpTags);

		// The last prop is the flag value we'll be updating, don't count it
		lpTags->cValues = ulNumOneOffIDs - 1;

		// If we're not removing the prop def stream, then don't count it
		if (!bRemovePropDef)
		{
			lpTags->cValues = lpTags->cValues - 1;
		}

		EC_MAPI(lpMessage->DeleteProps(
			lpTags,
			&lpProbArray));
		if (SUCCEEDED(hRes))
		{
			if (lpProbArray)
			{
				DebugPrint(DBGNamedProp, L"RemoveOneOff - DeleteProps problem array:\n%ws\n", ProblemArrayToString(lpProbArray).c_str());
			}

			SPropTagArray pTag = { 0 };
			ULONG cProp = 0;
			LPSPropValue lpCustomFlag = nullptr;

			// Grab dispidCustomFlag, the last tag in the array
			pTag.cValues = 1;
			pTag.aulPropTag[0] = CHANGE_PROP_TYPE(lpTags->aulPropTag[ulNumOneOffIDs - 1], PT_LONG);

			WC_MAPI(lpMessage->GetProps(
				&pTag,
				fMapiUnicode,
				&cProp,
				&lpCustomFlag));
			if (SUCCEEDED(hRes) && 1 == cProp && lpCustomFlag && PT_LONG == PROP_TYPE(lpCustomFlag->ulPropTag))
			{
				LPSPropProblemArray lpProbArray2 = nullptr;
				// Clear the INSP_ONEOFFFLAGS bits so OL doesn't look for the props we deleted
				lpCustomFlag->Value.l = lpCustomFlag->Value.l & ~INSP_ONEOFFFLAGS;
				if (bRemovePropDef)
				{
					lpCustomFlag->Value.l = lpCustomFlag->Value.l & ~INSP_PROPDEFINITION;
				}
				EC_MAPI(lpMessage->SetProps(
					1,
					lpCustomFlag,
					&lpProbArray2));
				if (S_OK == hRes && lpProbArray2)
				{
					DebugPrint(DBGNamedProp, L"RemoveOneOff - SetProps problem array:\n%ws\n", ProblemArrayToString(lpProbArray2).c_str());
				}

				MAPIFreeBuffer(lpProbArray2);
			}

			hRes = S_OK;

			EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

			if (SUCCEEDED(hRes))
			{
				DebugPrint(DBGNamedProp, L"RemoveOneOff - One-off properties removed.\n");
			}
		}

		MAPIFreeBuffer(lpProbArray);
	}

	MAPIFreeBuffer(lpTags);
	return hRes;
}

_Check_return_ HRESULT ResendMessages(_In_ LPMAPIFOLDER lpFolder, _In_ HWND hWnd)
{
	auto hRes = S_OK;
	LPMAPITABLE lpContentsTable = nullptr;
	LPSRowSet pRows = nullptr;
	ULONG i;

	// You define a SPropTagArray array here using the SizedSPropTagArray Macro
	// This enum will allows you to access portions of the array by a name instead of a number.
	// If more tags are added to the array, appropriate constants need to be added to the enum.
	enum
	{
		ePR_ENTRYID,
		NUM_COLS
	};
	// These tags represent the message information we would like to pick up
	static const SizedSPropTagArray(NUM_COLS, sptCols) =
	{
	NUM_COLS,
	PR_ENTRYID
	};

	if (!lpFolder) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpFolder->GetContentsTable(0, &lpContentsTable));

	if (lpContentsTable)
	{
		EC_MAPI(HrQueryAllRows(
			lpContentsTable,
			LPSPropTagArray(&sptCols),
			NULL, // restriction...we're not using this parameter
			NULL, // sort order...we're not using this parameter
			0,
			&pRows));

		if (pRows)
		{
			if (!FAILED(hRes)) for (i = 0; i < pRows->cRows; i++)
			{
				LPMESSAGE lpMessage = nullptr;

				hRes = S_OK;
				WC_H(CallOpenEntry(
					NULL,
					NULL,
					lpFolder,
					NULL,
					pRows->aRow[i].lpProps[ePR_ENTRYID].Value.bin.cb,
					reinterpret_cast<LPENTRYID>(pRows->aRow[i].lpProps[ePR_ENTRYID].Value.bin.lpb),
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
				if (lpMessage)
				{
					EC_H(ResendSingleMessage(lpFolder, lpMessage, hWnd));
					lpMessage->Release();
				}
			}
		}
	}

	if (pRows) FreeProws(pRows);
	if (lpContentsTable) lpContentsTable->Release();
	return hRes;
}

_Check_return_ HRESULT ResendSingleMessage(
	_In_ LPMAPIFOLDER lpFolder,
	_In_ LPSBinary MessageEID,
	_In_ HWND hWnd)
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;

	if (!lpFolder || !MessageEID) return MAPI_E_INVALID_PARAMETER;

	EC_H(CallOpenEntry(
		NULL,
		NULL,
		lpFolder,
		NULL,
		MessageEID->cb,
		reinterpret_cast<LPENTRYID>(MessageEID->lpb),
		NULL,
		MAPI_BEST_ACCESS,
		NULL,
		reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
	if (lpMessage)
	{
		EC_H(ResendSingleMessage(
			lpFolder,
			lpMessage,
			hWnd));
	}

	if (lpMessage) lpMessage->Release();
	return hRes;
}

_Check_return_ HRESULT ResendSingleMessage(
	_In_ LPMAPIFOLDER lpFolder,
	_In_ LPMESSAGE lpMessage,
	_In_ HWND hWnd)
{
	auto hRes = S_OK;
	auto hResRet = S_OK;
	LPATTACH lpAttach = nullptr;
	LPMESSAGE lpAttachMsg = nullptr;
	LPMAPITABLE lpAttachTable = nullptr;
	LPSRowSet pRows = nullptr;
	LPMESSAGE lpNewMessage = nullptr;
	LPSPropTagArray lpsMessageTags = nullptr;
	LPSPropProblemArray lpsPropProbs = nullptr;
	SPropValue sProp = { 0 };

	enum
	{
		atPR_ATTACH_METHOD,
		atPR_ATTACH_NUM,
		atPR_DISPLAY_NAME,
		atNUM_COLS
	};

	static const SizedSPropTagArray(atNUM_COLS, atCols) =
	{
	atNUM_COLS,
	PR_ATTACH_METHOD,
	PR_ATTACH_NUM,
	PR_DISPLAY_NAME_W
	};

	static const SizedSPropTagArray(2, atObjs) =
	{
	2,
	PR_MESSAGE_RECIPIENTS,
	PR_MESSAGE_ATTACHMENTS
	};

	if (!lpMessage || !lpFolder) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"ResendSingleMessage: Checking message for embedded messages\n");

	EC_MAPI(lpMessage->GetAttachmentTable(
		NULL,
		&lpAttachTable));

	if (lpAttachTable)
	{
		EC_MAPI(lpAttachTable->SetColumns(LPSPropTagArray(&atCols), TBL_BATCH));

		// Now we iterate through each of the attachments
		if (!FAILED(hRes)) for (;;)
		{
			// Remember the first error code we hit so it will bubble up
			if (FAILED(hRes) && SUCCEEDED(hResRet)) hResRet = hRes;
			hRes = S_OK;
			if (pRows) FreeProws(pRows);
			pRows = nullptr;
			EC_MAPI(lpAttachTable->QueryRows(
				1,
				NULL,
				&pRows));
			if (FAILED(hRes)) break;
			if (pRows && (!pRows || pRows->cRows)) break;

			if (ATTACH_EMBEDDED_MSG == pRows->aRow->lpProps[atPR_ATTACH_METHOD].Value.l)
			{
				DebugPrint(DBGGeneric, L"Found an embedded message to resend.\n");

				if (lpAttach) lpAttach->Release();
				lpAttach = nullptr;
				EC_MAPI(lpMessage->OpenAttach(
					pRows->aRow->lpProps[atPR_ATTACH_NUM].Value.l,
					NULL,
					MAPI_BEST_ACCESS,
					static_cast<LPATTACH*>(&lpAttach)));
				if (!lpAttach) continue;

				if (lpAttachMsg) lpAttachMsg->Release();
				lpAttachMsg = nullptr;
				EC_MAPI(lpAttach->OpenProperty(
					PR_ATTACH_DATA_OBJ,
					const_cast<LPIID>(&IID_IMessage),
					0,
					MAPI_MODIFY,
					reinterpret_cast<LPUNKNOWN *>(&lpAttachMsg)));
				if (MAPI_E_INTERFACE_NOT_SUPPORTED == hRes)
				{
					CHECKHRESMSG(hRes, IDS_ATTNOTEMBEDDEDMSG);
					continue;
				}

				if (FAILED(hRes)) continue;

				DebugPrint(DBGGeneric, L"Message opened.\n");

				if (CheckStringProp(&pRows->aRow->lpProps[atPR_DISPLAY_NAME], PT_UNICODE))
				{
					DebugPrint(DBGGeneric, L"Resending \"%ws\"\n", pRows->aRow->lpProps[atPR_DISPLAY_NAME].Value.lpszW);
				}

				DebugPrint(DBGGeneric, L"Creating new message.\n");
				if (lpNewMessage) lpNewMessage->Release();
				lpNewMessage = nullptr;
				EC_MAPI(lpFolder->CreateMessage(NULL, 0, &lpNewMessage));
				if (!lpNewMessage) continue;

				EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));

				// Copy all the transmission properties
				DebugPrint(DBGGeneric, L"Getting list of properties.\n");
				MAPIFreeBuffer(lpsMessageTags);
				lpsMessageTags = nullptr;
				EC_MAPI(lpAttachMsg->GetPropList(0, &lpsMessageTags));
				if (!lpsMessageTags) continue;

				DebugPrint(DBGGeneric, L"Copying properties to new message.\n");
				if (!FAILED(hRes)) for (ULONG ulProp = 0; ulProp < lpsMessageTags->cValues; ulProp++)
				{
					hRes = S_OK;
					// it would probably be quicker to use this loop to construct an array of properties
					// we desire to copy, and then pass that array to GetProps and then SetProps
					if (FIsTransmittable(lpsMessageTags->aulPropTag[ulProp]))
					{
						LPSPropValue lpProp = nullptr;
						DebugPrint(DBGGeneric, L"Copying 0x%08X\n", lpsMessageTags->aulPropTag[ulProp]);
						WC_MAPI(HrGetOneProp(lpAttachMsg, lpsMessageTags->aulPropTag[ulProp], &lpProp));

						WC_MAPI(HrSetOneProp(lpNewMessage, lpProp));

						MAPIFreeBuffer(lpProp);
					}
				}

				EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));

				DebugPrint(DBGGeneric, L"Copying recipients and attachments to new message.\n");

				LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIProp::CopyProps", hWnd); // STRING_OK

				EC_MAPI(lpAttachMsg->CopyProps(
					LPSPropTagArray(&atObjs),
					lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
					lpProgress,
					&IID_IMessage,
					lpNewMessage,
					lpProgress ? MAPI_DIALOG : 0,
					&lpsPropProbs));
				if (lpsPropProbs)
				{
					EC_PROBLEMARRAY(lpsPropProbs);
					MAPIFreeBuffer(lpsPropProbs);
					lpsPropProbs = nullptr;
					continue;
				}

				if (lpProgress) lpProgress->Release();

				sProp.dwAlignPad = 0;
				sProp.ulPropTag = PR_DELETE_AFTER_SUBMIT;
				sProp.Value.b = true;

				DebugPrint(DBGGeneric, L"Setting PR_DELETE_AFTER_SUBMIT to true.\n");
				EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));

				SPropTagArray sPropTagArray = { 0 };

				sPropTagArray.cValues = 1;
				sPropTagArray.aulPropTag[0] = PR_SENTMAIL_ENTRYID;

				DebugPrint(DBGGeneric, L"Deleting PR_SENTMAIL_ENTRYID\n");
				EC_MAPI(lpNewMessage->DeleteProps(&sPropTagArray, NULL));

				EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));

				DebugPrint(DBGGeneric, L"Submitting new message.\n");
				EC_MAPI(lpNewMessage->SubmitMessage(0));
			}
			else
			{
				DebugPrint(DBGGeneric, L"Attachment is not an embedded message.\n");
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
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	enum
	{
		eidPR_ENTRYID,
		eidNUM_COLS
	};

	static const SizedSPropTagArray(eidNUM_COLS, eidCols) =
	{
	eidNUM_COLS,
	PR_ENTRYID,
	};

	if (!lpMDB || !lpMAPIFolder) return MAPI_E_INVALID_PARAMETER;

	// We pass through this code twice, once for regular contents, once for associated contents
	if (!FAILED(hRes)) for (auto i = 0; i <= 1; i++)
	{
		hRes = S_OK;
		auto ulFlags = (1 == i ? MAPI_ASSOCIATED : NULL) |
			fMapiUnicode;

		if (lpContentsTable) lpContentsTable->Release();
		lpContentsTable = nullptr;
		// Get the table of contents of the folder
		EC_MAPI(lpMAPIFolder->GetContentsTable(
			ulFlags,
			&lpContentsTable));

		if (SUCCEEDED(hRes) && lpContentsTable)
		{
			EC_MAPI(lpContentsTable->SetColumns(
				LPSPropTagArray(&eidCols),
				TBL_BATCH));

			// go to the first row
			EC_MAPI(lpContentsTable->SeekRow(
				BOOKMARK_BEGINNING,
				0,
				NULL));
			hRes = S_OK; // don't let failure here fail the whole op

			// get rows and delete PR_NT_SECURITY_DESCRIPTOR
			if (!FAILED(hRes)) for (;;)
			{
				hRes = S_OK;
				if (pRows) FreeProws(pRows);
				pRows = nullptr;
				// Pull back a sizable block of rows to modify
				EC_MAPI(lpContentsTable->QueryRows(
					200,
					NULL,
					&pRows));
				if (FAILED(hRes) || !pRows || !pRows->cRows) break;

				for (ULONG iCurPropRow = 0; iCurPropRow < pRows->cRows; iCurPropRow++)
				{
					hRes = S_OK;
					if (lpMessage) lpMessage->Release();
					lpMessage = nullptr;

					WC_H(CallOpenEntry(
						lpMDB,
						NULL,
						NULL,
						NULL,
						pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID].Value.bin.cb,
						reinterpret_cast<LPENTRYID>(pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID].Value.bin.lpb),
						NULL,
						MAPI_BEST_ACCESS,
						NULL,
						reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
					if (FAILED(hRes))
					{
						hResOverall = hRes;
						continue;
					}

					WC_H(DeleteProperty(lpMessage, PR_NT_SECURITY_DESCRIPTOR));
					if (FAILED(hRes))
					{
						hResOverall = hRes;
						continue;
					}

					iItemCount++;
				}
			}

			DebugPrint(DBGGeneric, L"ResetPermissionsOnItems reset permissions on %u items\n", iItemCount);
		}
	}

	if (pRows) FreeProws(pRows);
	if (lpMessage) lpMessage->Release();
	if (lpContentsTable) lpContentsTable->Release();
	if (S_OK != hResOverall) return hResOverall;
	return hRes;
}

// This function creates a new message based in m_lpContainer
// Then sends the message
_Check_return_ HRESULT SendTestMessage(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMAPIFOLDER lpFolder,
	_In_ wstring szRecipient,
	_In_ wstring szBody,
	_In_ wstring& szSubject,
	_In_ wstring szClass)
{
	auto hRes = S_OK;
	LPMESSAGE lpNewMessage = nullptr;

	if (!lpMAPISession || !lpFolder) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpFolder->CreateMessage(
		NULL, // default interface
		0, // flags
		&lpNewMessage));

	if (lpNewMessage)
	{
		SPropValue sProp;

		sProp.dwAlignPad = 0;
		sProp.ulPropTag = PR_DELETE_AFTER_SUBMIT;
		sProp.Value.b = true;

		DebugPrint(DBGGeneric, L"Setting PR_DELETE_AFTER_SUBMIT to true.\n");
		EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));

		sProp.dwAlignPad = 0;
		sProp.ulPropTag = PR_BODY_W;
		sProp.Value.lpszW = const_cast<LPWSTR>(szBody.c_str());

		DebugPrint(DBGGeneric, L"Setting PR_BODY to %ws.\n", szBody.c_str());
		EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));

		sProp.dwAlignPad = 0;
		sProp.ulPropTag = PR_SUBJECT_W;
		sProp.Value.lpszW = const_cast<LPWSTR>(szSubject.c_str());

		DebugPrint(DBGGeneric, L"Setting PR_SUBJECT to %ws.\n", szSubject.c_str());
		EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));

		sProp.dwAlignPad = 0;
		sProp.ulPropTag = PR_MESSAGE_CLASS_W;
		sProp.Value.lpszW = const_cast<LPWSTR>(szClass.c_str());

		DebugPrint(DBGGeneric, L"Setting PR_MESSAGE_CLASS to %ws.\n", szSubject.c_str());
		EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));

		SPropTagArray sPropTagArray;

		sPropTagArray.cValues = 1;
		sPropTagArray.aulPropTag[0] = PR_SENTMAIL_ENTRYID;

		DebugPrint(DBGGeneric, L"Deleting PR_SENTMAIL_ENTRYID\n");
		EC_MAPI(lpNewMessage->DeleteProps(&sPropTagArray, NULL));

		DebugPrint(DBGGeneric, L"Adding recipient: %ws.\n", szRecipient.c_str());
		EC_H(AddRecipient(
			lpMAPISession,
			lpNewMessage,
			szRecipient,
			MAPI_TO));

		DebugPrint(DBGGeneric, L"Submitting message\n");
		EC_MAPI(lpNewMessage->SubmitMessage(NULL));
	}

	if (lpNewMessage) lpNewMessage->Release();
	return hRes;
}

// Declaration missing from headers
STDAPI_(HRESULT) WrapCompressedRTFStreamEx(
	IStream* pCompressedRTFStream,
	CONST RTF_WCSINFO * pWCSInfo,
	IStream** ppUncompressedRTFStream,
	RTF_WCSRETINFO * pRetInfo);

_Check_return_ HRESULT WrapStreamForRTF(
	_In_ LPSTREAM lpCompressedRTFStream,
	bool bUseWrapEx,
	ULONG ulFlags,
	ULONG ulInCodePage,
	ULONG ulOutCodePage,
	_Deref_out_ LPSTREAM* lpUncompressedRTFStream,
	_Out_opt_ ULONG* pulStreamFlags)
{
	if (!lpCompressedRTFStream || !lpUncompressedRTFStream) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;

	if (!bUseWrapEx)
	{
		WC_MAPI(WrapCompressedRTFStream(
			lpCompressedRTFStream,
			ulFlags,
			lpUncompressedRTFStream));
	}
	else
	{
		RTF_WCSINFO wcsinfo = { 0 };
		RTF_WCSRETINFO retinfo = { 0 };

		retinfo.size = sizeof(RTF_WCSRETINFO);

		wcsinfo.size = sizeof(RTF_WCSINFO);
		wcsinfo.ulFlags = ulFlags;
		wcsinfo.ulInCodePage = ulInCodePage; // Get ulCodePage from PR_INTERNET_CPID on the IMessage
		wcsinfo.ulOutCodePage = ulOutCodePage; // Desired code page for return

		WC_MAPI(WrapCompressedRTFStreamEx(
			lpCompressedRTFStream,
			&wcsinfo,
			lpUncompressedRTFStream,
			&retinfo));
		if (pulStreamFlags) *pulStreamFlags = retinfo.ulStreamFlags;
	}

	return hRes;
}

_Check_return_ HRESULT CopyNamedProps(_In_ LPMAPIPROP lpSource, _In_ LPGUID lpPropSetGUID, bool bDoMove, bool bDoNoReplace, _In_ LPMAPIPROP lpTarget, _In_ HWND hWnd)
{
	if (!lpSource || !lpTarget) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;
	LPSPropTagArray lpPropTags = nullptr;

	EC_H(GetNamedPropsByGUID(lpSource, lpPropSetGUID, &lpPropTags));

	if (!FAILED(hRes) && lpPropTags)
	{
		LPSPropProblemArray lpProblems = nullptr;
		ULONG ulFlags = 0;
		if (bDoMove) ulFlags |= MAPI_MOVE;
		if (bDoNoReplace) ulFlags |= MAPI_NOREPLACE;

		LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIProp::CopyProps", hWnd); // STRING_OK

		if (lpProgress)
			ulFlags |= MAPI_DIALOG;

		EC_MAPI(lpSource->CopyProps(lpPropTags,
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

_Check_return_ HRESULT GetNamedPropsByGUID(_In_ LPMAPIPROP lpSource, _In_ LPGUID lpPropSetGUID, _Deref_out_ LPSPropTagArray* lpOutArray)
{
	if (!lpSource || !lpPropSetGUID || lpOutArray)
		return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;
	LPSPropTagArray lpAllProps = nullptr;

	*lpOutArray = nullptr;

	WC_MAPI(lpSource->GetPropList(0, &lpAllProps));

	if (S_OK == hRes && lpAllProps)
	{
		ULONG cProps = 0;
		LPMAPINAMEID* lppNameIDs = nullptr;

		WC_H(GetNamesFromIDs(lpSource,
			&lpAllProps,
			NULL,
			0,
			&cProps,
			&lppNameIDs));

		if (S_OK == hRes && lppNameIDs)
		{
			ULONG ulNumProps = 0; // count of props that match our guid
			for (ULONG i = 0; i < cProps; i++)
			{
				if (PROP_ID(lpAllProps->aulPropTag[i]) > 0x7FFF &&
					lppNameIDs[i] &&
					!memcmp(lppNameIDs[i]->lpguid, lpPropSetGUID, sizeof(GUID)))
				{
					ulNumProps++;
				}
			}

			LPSPropTagArray lpFilteredProps = nullptr;

			WC_H(MAPIAllocateBuffer(CbNewSPropTagArray(ulNumProps),
				reinterpret_cast<LPVOID*>(&lpFilteredProps)));

			if (S_OK == hRes && lpFilteredProps)
			{
				lpFilteredProps->cValues = 0;

				for (ULONG i = 0; i < cProps; i++)
				{
					if (PROP_ID(lpAllProps->aulPropTag[i]) > 0x7FFF &&
						lppNameIDs[i] &&
						!memcmp(lppNameIDs[i]->lpguid, lpPropSetGUID, sizeof(GUID)))
					{
						lpFilteredProps->aulPropTag[lpFilteredProps->cValues] = lpAllProps->aulPropTag[i];
						lpFilteredProps->cValues++;
					}
				}

				*lpOutArray = lpFilteredProps;
			}
		}

		MAPIFreeBuffer(lppNameIDs);
	}

	MAPIFreeBuffer(lpAllProps);
	return hRes;
}

_Check_return_ bool CheckStringProp(_In_opt_ LPSPropValue lpProp, ULONG ulPropType)
{
	if (PT_STRING8 != ulPropType && PT_UNICODE != ulPropType)
	{
		DebugPrint(DBGGeneric, L"CheckStringProp: Called with invalid ulPropType of 0x%X\n", ulPropType);
		return false;
	}

	if (!lpProp)
	{
		DebugPrint(DBGGeneric, L"CheckStringProp: lpProp is NULL\n");
		return false;
	}

	if (PT_ERROR == PROP_TYPE(lpProp->ulPropTag))
	{
		DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->ulPropTag is of type PT_ERROR\n");
		return false;
	}

	if (ulPropType != PROP_TYPE(lpProp->ulPropTag))
	{
		DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->ulPropTag is not of type 0x%X\n", ulPropType);
		return false;
	}

	if (NULL == lpProp->Value.LPSZ)
	{
		DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->Value.LPSZ is NULL\n");
		return false;
	}

	if (PT_STRING8 == ulPropType && NULL == lpProp->Value.lpszA[0])
	{
		DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->Value.lpszA[0] is NULL\n");
		return false;
	}

	if (PT_UNICODE == ulPropType && NULL == lpProp->Value.lpszW[0])
	{
		DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->Value.lpszW[0] is NULL\n");
		return false;
	}

	return true;
}

_Check_return_ DWORD ComputeStoreHash(ULONG cbStoreEID, _In_count_(cbStoreEID) LPBYTE pbStoreEID, _In_opt_z_ LPCSTR pszFileName, _In_opt_z_ LPCWSTR pwzFileName, bool bPublicStore)
{
	DWORD dwHash = 0;

	if (!cbStoreEID || !pbStoreEID) return dwHash;
	// We shouldn't see both of these at the same time.
	if (pszFileName && pwzFileName) return dwHash;

	// Get the Store Entry ID
	// pbStoreEID is a pointer to the Entry ID
	// cbStoreEID is the size in bytes of the Entry ID
	auto pdw = reinterpret_cast<LPDWORD>(pbStoreEID);
	auto cdw = cbStoreEID / sizeof(DWORD);

	for (ULONG i = 0; i < cdw; i++)
	{
		dwHash = (dwHash << 5) + dwHash + *pdw++;
	}

	auto pb = reinterpret_cast<LPBYTE>(pdw);
	auto cb = cbStoreEID % sizeof(DWORD);

	for (ULONG i = 0; i < cb; i++)
	{
		dwHash = (dwHash << 5) + dwHash + *pb++;
	}

	if (bPublicStore)
	{
		DebugPrint(DBGGeneric, L"ComputeStoreHash, hash (before adding .PUB) = 0x%08X\n", dwHash);
		// augment to make sure it is unique else could be same as the private store
		dwHash = (dwHash << 5) + dwHash + 0x2E505542; // this is '.PUB'
	}

	if (pwzFileName || pszFileName)
		DebugPrint(DBGGeneric, L"ComputeStoreHash, hash (before adding path) = 0x%08X\n", dwHash);

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
		DebugPrint(DBGGeneric, L"ComputeStoreHash, hash (after adding path) = 0x%08X\n", dwHash);

	// dwHash now contains the hash to be used. It should be written in hex when building a URL.
	return dwHash;
}

const WORD kwBaseOffset = 0xAC00; // Hangul char range (AC00-D7AF)
// Allocates with new, free with delete[]
_Check_return_ wstring EncodeID(ULONG cbEID, _In_ LPENTRYID rgbID)
{
	auto pbSrc = reinterpret_cast<LPBYTE>(rgbID);
	wstring wzIDEncoded;

	// rgbID is the item Entry ID or the attachment ID
	// cbID is the size in bytes of rgbID
	for (ULONG i = 0; i < cbEID; i++, pbSrc++)
	{
		wzIDEncoded += static_cast<WCHAR>(*pbSrc + kwBaseOffset);
	}

	// pwzIDEncoded now contains the entry ID encoded.
	return wzIDEncoded;
}

_Check_return_ wstring DecodeID(ULONG cbBuffer, _In_count_(cbBuffer) LPBYTE lpbBuffer)
{
	if (cbBuffer % 2) return emptystring;

	auto cbDecodedBuffer = cbBuffer / 2;
	// Allocate memory for lpDecoded
	auto lpDecoded = new BYTE[cbDecodedBuffer];
	if (!lpDecoded) return emptystring;

	// Subtract kwBaseOffset from every character and place result in lpDecoded
	auto lpwzSrc = reinterpret_cast<LPWSTR>(lpbBuffer);
	auto lpDst = lpDecoded;
	for (ULONG i = 0; i < cbDecodedBuffer; i++, lpwzSrc++, lpDst++)
	{
		*lpDst = static_cast<BYTE>(*lpwzSrc - kwBaseOffset);
	}

	auto szBin = BinToHexString(
		lpDecoded,
		cbDecodedBuffer,
		true);
	delete[] lpDecoded;
	return szBin;
}

HRESULT HrEmsmdbUIDFromStore(_In_ LPMAPISESSION pmsess,
	_In_ MAPIUID const * const puidService,
	_Out_opt_ MAPIUID* pEmsmdbUID)
{
	if (!puidService) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;

	SRestriction mres = { 0 };
	SPropValue mval = { 0 };
	SRowSet* pRows = nullptr;
	LPSERVICEADMIN spSvcAdmin = nullptr;
	LPMAPITABLE spmtab = nullptr;

	enum
	{
		eEntryID = 0,
		eSectionUid,
		eMax
	};
	static const SizedSPropTagArray(eMax, tagaCols) =
	{
	eMax,
	{
	PR_ENTRYID,
	PR_EMSMDB_SECTION_UID,
	}
	};

	EC_MAPI(pmsess->AdminServices(0, static_cast<LPSERVICEADMIN*>(&spSvcAdmin)));
	if (spSvcAdmin)
	{
		EC_MAPI(spSvcAdmin->GetMsgServiceTable(0, &spmtab));
		if (spmtab)
		{
			EC_MAPI(spmtab->SetColumns(LPSPropTagArray(&tagaCols), TBL_BATCH));

			mres.rt = RES_PROPERTY;
			mres.res.resProperty.relop = RELOP_EQ;
			mres.res.resProperty.ulPropTag = PR_SERVICE_UID;
			mres.res.resProperty.lpProp = &mval;
			mval.ulPropTag = PR_SERVICE_UID;
			mval.Value.bin.cb = sizeof*puidService;
			mval.Value.bin.lpb = LPBYTE(puidService);

			EC_MAPI(spmtab->Restrict(&mres, 0));
			EC_MAPI(spmtab->QueryRows(10, 0, &pRows));

			if (SUCCEEDED(hRes) && pRows && pRows->cRows)
			{
				auto pRow = &pRows->aRow[0];

				if (pEmsmdbUID && pRow)
				{
					if (PR_EMSMDB_SECTION_UID == pRow->lpProps[eSectionUid].ulPropTag &&
						pRow->lpProps[eSectionUid].Value.bin.cb == sizeof*pEmsmdbUID)
					{
						memcpy(pEmsmdbUID, pRow->lpProps[eSectionUid].Value.bin.lpb, sizeof*pEmsmdbUID);
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

bool FExchangePrivateStore(_In_ LPMAPIUID lpmapiuid)
{
	if (!lpmapiuid) return false;
	return IsEqualMAPIUID(lpmapiuid, (LPMAPIUID)pbExchangeProviderPrimaryUserGuid);
}

bool FExchangePublicStore(_In_ LPMAPIUID lpmapiuid)
{
	if (!lpmapiuid) return false;
	return IsEqualMAPIUID(lpmapiuid, (LPMAPIUID)pbExchangeProviderPublicGuid);
}

STDMETHODIMP GetEntryIDFromMDB(LPMDB lpMDB, ULONG ulPropTag, _Out_opt_ ULONG* lpcbeid, _Deref_out_opt_ LPENTRYID* lppeid)
{
	if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;
	LPSPropValue lpEIDProp = nullptr;

	WC_MAPI(HrGetOneProp(lpMDB, ulPropTag, &lpEIDProp));

	if (SUCCEEDED(hRes) && lpEIDProp)
	{
		WC_H(MAPIAllocateBuffer(lpEIDProp->Value.bin.cb, reinterpret_cast<LPVOID*>(lppeid)));
		if (SUCCEEDED(hRes))
		{
			*lpcbeid = lpEIDProp->Value.bin.cb;
			CopyMemory(*lppeid, lpEIDProp->Value.bin.lpb, *lpcbeid);
		}
	}

	MAPIFreeBuffer(lpEIDProp);
	return hRes;
}

STDMETHODIMP GetMVEntryIDFromInboxByIndex(LPMDB lpMDB, ULONG ulPropTag, ULONG ulIndex, _Out_opt_ ULONG* lpcbeid, _Deref_out_opt_ LPENTRYID* lppeid)
{
	if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;
	LPMAPIFOLDER lpInbox = nullptr;

	WC_H(GetInbox(lpMDB, &lpInbox));

	if (SUCCEEDED(hRes) && lpInbox)
	{
		LPSPropValue lpEIDProp = nullptr;
		WC_MAPI(HrGetOneProp(lpInbox, ulPropTag, &lpEIDProp));
		if (SUCCEEDED(hRes) &&
			lpEIDProp &&
			PT_MV_BINARY == PROP_TYPE(lpEIDProp->ulPropTag) &&
			ulIndex < lpEIDProp->Value.MVbin.cValues &&
			lpEIDProp->Value.MVbin.lpbin[ulIndex].cb > 0)
		{
			WC_H(MAPIAllocateBuffer(lpEIDProp->Value.MVbin.lpbin[ulIndex].cb, reinterpret_cast<LPVOID*>(lppeid)));
			if (SUCCEEDED(hRes))
			{
				*lpcbeid = lpEIDProp->Value.MVbin.lpbin[ulIndex].cb;
				CopyMemory(*lppeid, lpEIDProp->Value.MVbin.lpbin[ulIndex].lpb, *lpcbeid);
			}
		}

		MAPIFreeBuffer(lpEIDProp);
	}
	if (lpInbox) lpInbox->Release();

	return hRes;
}

STDMETHODIMP GetDefaultFolderEID(
	_In_ ULONG ulFolder,
	_In_ LPMDB lpMDB,
	_Out_opt_ ULONG* lpcbeid,
	_Deref_out_opt_ LPENTRYID* lppeid)
{
	auto hRes = S_OK;

	if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;

	switch (ulFolder)
	{
	case DEFAULT_CALENDAR:
		hRes = GetSpecialFolderEID(lpMDB, PR_IPM_APPOINTMENT_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_CONTACTS:
		hRes = GetSpecialFolderEID(lpMDB, PR_IPM_CONTACT_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_JOURNAL:
		hRes = GetSpecialFolderEID(lpMDB, PR_IPM_JOURNAL_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_NOTES:
		hRes = GetSpecialFolderEID(lpMDB, PR_IPM_NOTE_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_TASKS:
		hRes = GetSpecialFolderEID(lpMDB, PR_IPM_TASK_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_REMINDERS:
		hRes = GetSpecialFolderEID(lpMDB, PR_REM_ONLINE_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_DRAFTS:
		hRes = GetSpecialFolderEID(lpMDB, PR_IPM_DRAFTS_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_SENTITEMS:
		hRes = GetEntryIDFromMDB(lpMDB, PR_IPM_SENTMAIL_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_OUTBOX:
		hRes = GetEntryIDFromMDB(lpMDB, PR_IPM_OUTBOX_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_DELETEDITEMS:
		hRes = GetEntryIDFromMDB(lpMDB, PR_IPM_WASTEBASKET_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_FINDER:
		hRes = GetEntryIDFromMDB(lpMDB, PR_FINDER_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_IPM_SUBTREE:
		hRes = GetEntryIDFromMDB(lpMDB, PR_IPM_SUBTREE_ENTRYID, lpcbeid, lppeid);
		break;
	case DEFAULT_INBOX:
		hRes = GetInbox(lpMDB, lpcbeid, lppeid);
		break;
	case DEFAULT_LOCALFREEBUSY:
		hRes = GetMVEntryIDFromInboxByIndex(lpMDB, PR_FREEBUSY_ENTRYIDS, 3, lpcbeid, lppeid);
		break;
	case DEFAULT_CONFLICTS:
		hRes = GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 0, lpcbeid, lppeid);
		break;
	case DEFAULT_SYNCISSUES:
		hRes = GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 1, lpcbeid, lppeid);
		break;
	case DEFAULT_LOCALFAILURES:
		hRes = GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 2, lpcbeid, lppeid);
		break;
	case DEFAULT_SERVERFAILURES:
		hRes = GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 3, lpcbeid, lppeid);
		break;
	case DEFAULT_JUNKMAIL:
		hRes = GetMVEntryIDFromInboxByIndex(lpMDB, PR_ADDITIONAL_REN_ENTRYIDS, 4, lpcbeid, lppeid);
		break;
	default:
		hRes = MAPI_E_INVALID_PARAMETER;
	}

	return hRes;
}

STDMETHODIMP OpenDefaultFolder(_In_ ULONG ulFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER *lpFolder)
{
	auto hRes = S_OK;

	if (!lpMDB || !lpFolder) return MAPI_E_INVALID_PARAMETER;

	*lpFolder = nullptr;
	ULONG cb = 0;
	LPENTRYID lpeid = nullptr;

	WC_H(GetDefaultFolderEID(ulFolder, lpMDB, &cb, &lpeid));
	if (SUCCEEDED(hRes))
	{
		LPMAPIFOLDER lpTemp = nullptr;
		WC_H(CallOpenEntry(
			lpMDB,
			NULL,
			NULL,
			NULL,
			cb,
			lpeid,
			NULL,
			MAPI_BEST_ACCESS,
			NULL,
			reinterpret_cast<LPUNKNOWN*>(&lpTemp)));
		if (SUCCEEDED(hRes) && lpTemp)
		{
			*lpFolder = lpTemp;
		}
		else if (lpTemp)
		{
			lpTemp->Release();
		}
	}

	MAPIFreeBuffer(lpeid);
	return hRes;
}


ULONG g_DisplayNameProps[] =
{
 PR_DISPLAY_NAME_W,
 CHANGE_PROP_TYPE(PR_MAILBOX_OWNER_NAME, PT_UNICODE),
 PR_SUBJECT_W,
};


wstring GetTitle(LPMAPIPROP lpMAPIProp)
{
	auto hRes = S_OK;
	wstring szTitle;
	LPSPropValue lpProp = nullptr;
	auto bFoundName = false;

	if (!lpMAPIProp) return szTitle;

	// Get a property for the title bar
	for (ULONG i = 0; !bFoundName && i < _countof(g_DisplayNameProps); i++)
	{
		hRes = S_OK;
		WC_MAPI(HrGetOneProp(
			lpMAPIProp,
			g_DisplayNameProps[i],
			&lpProp));

		if (lpProp)
		{
			if (CheckStringProp(lpProp, PT_UNICODE))
			{
				szTitle = lpProp->Value.lpszW;
				bFoundName = true;
			}

			MAPIFreeBuffer(lpProp);
		}
	}

	if (!bFoundName)
	{
		szTitle = loadstring(IDS_DISPLAYNAMENOTFOUND);
	}

	return szTitle;
}

bool UnwrapContactEntryID(_In_ ULONG cbIn, _In_ LPBYTE lpbIn, _Out_ ULONG* lpcbOut, _Out_ LPBYTE* lppbOut)
{
	if (lpcbOut) *lpcbOut = 0;
	if (lppbOut) *lppbOut = nullptr;

	if (cbIn < sizeof(DIR_ENTRYID)) return false;
	if (!lpcbOut || !lppbOut || !lpbIn) return false;

	auto lpContabEID = reinterpret_cast<LPCONTAB_ENTRYID>(lpbIn);

	switch (lpContabEID->ulType)
	{
	case CONTAB_USER:
	case CONTAB_DISTLIST:
		if (cbIn >= sizeof(CONTAB_ENTRYID) && lpContabEID->cbeid && lpContabEID->abeid)
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

#ifndef MRMAPI
// Takes a tag array (and optional MAPIProp) and displays UI prompting to build an exclusion array
// Must be freed with MAPIFreeBuffer
LPSPropTagArray GetExcludedTags(_In_opt_ LPSPropTagArray lpTagArray, _In_opt_ LPMAPIPROP lpProp, bool bIsAB)
{
	auto hRes = S_OK;

	CTagArrayEditor TagEditor(
		nullptr,
		IDS_TAGSTOEXCLUDE,
		IDS_TAGSTOEXCLUDEPROMPT,
		nullptr,
		lpTagArray,
		bIsAB,
		lpProp);
	WC_H(TagEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		return TagEditor.DetachModifiedTagArray();
	}

	return nullptr;
}
#endif

// Performs CopyTo operation from source to destination, optionally prompting for exclusions
// Does not save changes - caller should do this.
HRESULT CopyTo(HWND hWnd, _In_ LPMAPIPROP lpSource, _In_ LPMAPIPROP lpDest, LPCGUID lpGUID, _In_opt_ LPSPropTagArray lpTagArray, bool bIsAB, bool bAllowUI)
{
	auto hRes = S_OK;
	if (!lpSource || !lpDest) return MAPI_E_INVALID_PARAMETER;

	LPSPropProblemArray lpProblems = nullptr;
	auto lpExcludedTags = lpTagArray;
	LPSPropTagArray lpUITags = nullptr;
	LPMAPIPROGRESS lpProgress = nullptr;
	auto lpGUIDLocal = lpGUID;
	ULONG ulCopyFlags = 0;

#ifdef MRMAPI
	bAllowUI, bIsAB; // Unused parameters in MRMAPI
#endif

#ifndef MRMAPI
	if (bAllowUI)
	{
		CEditor MyData(
			nullptr,
			IDS_COPYTO,
			IDS_COPYPASTEPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_INTERFACE, GUIDToStringAndName(lpGUID), false));
		MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_FLAGS, false));
		MyData.SetHex(1, MAPI_DIALOG);

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			auto MyGUID = StringToGUID(MyData.GetStringW(0));
			lpGUIDLocal = &MyGUID;
			ulCopyFlags = MyData.GetHex(1);
			if (hWnd)
			{
				lpProgress = GetMAPIProgress(L"CopyTo", hWnd); // STRING_OK
				if (lpProgress)
					ulCopyFlags |= MAPI_DIALOG;
			}

			lpUITags = GetExcludedTags(lpTagArray, lpSource, bIsAB);
			if (lpUITags)
			{
				lpExcludedTags = lpUITags;
			}
		}
	}
#endif

	if (S_OK == hRes)
	{
		EC_MAPI(lpSource->CopyTo(
			0,
			NULL,
			lpExcludedTags,
			lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL, // UI param
			lpProgress, // progress
			lpGUIDLocal,
			lpDest,
			ulCopyFlags, // flags
			&lpProblems));
	}

	MAPIFreeBuffer(lpUITags);
	if (lpProgress) lpProgress->Release();

	if (lpProblems)
	{
		EC_PROBLEMARRAY(lpProblems);
		MAPIFreeBuffer(lpProblems);
	}

	return hRes;
}