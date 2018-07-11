// Collection of useful MAPI functions

#include <StdAfx.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/String.h>
#include <MAPI/MAPIABFunctions.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>
#include <MAPI/MAPIProgress.h>
#include <Interpret/Guids.h>
#include <MAPI/Cache/NamedPropCache.h>
#include <Interpret/SmartView/SmartView.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <MAPI/MapiMemory.h>

namespace mapi
{
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

		if (registry::RegKeys[registry::regKeyMAPI_NO_CACHE].ulCurDWORD)
		{
			ulFlags |= MAPI_NO_CACHE;
		}

		// in case we need to retry without MAPI_NO_CACHE - do not add MAPI_NO_CACHE to ulFlags after this point
		if (MAPI_NO_CACHE & ulFlags) ulNoCacheFlags = ulFlags & ~MAPI_NO_CACHE;

		if (lpInterface && fIsSet(DBGGeneric))
		{
			auto szGuid = guid::GUIDToStringAndName(lpInterface);
			output::DebugPrint(DBGGeneric, L"CallOpenEntry: OpenEntry asking for %ws\n", szGuid.c_str());
		}

		if (lpMDB)
		{
			output::DebugPrint(DBGGeneric, L"CallOpenEntry: Calling OpenEntry on MDB with ulFlags = 0x%X\n", ulFlags);
			hRes = WC_MAPI(lpMDB->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulFlags, &ulObjType, &lpUnk));
			if (hRes == MAPI_E_UNKNOWN_FLAGS && ulNoCacheFlags)
			{
				output::DebugPrint(
					DBGGeneric,
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
			output::DebugPrint(DBGGeneric, L"CallOpenEntry: Calling OpenEntry on AB with ulFlags = 0x%X\n", ulFlags);
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
					DBGGeneric,
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
				DBGGeneric, L"CallOpenEntry: Calling OpenEntry on Container with ulFlags = 0x%X\n", ulFlags);
			hRes = WC_MAPI(lpContainer->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulFlags, &ulObjType, &lpUnk));
			if (hRes == MAPI_E_UNKNOWN_FLAGS && ulNoCacheFlags)
			{
				output::DebugPrint(
					DBGGeneric,
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
				DBGGeneric, L"CallOpenEntry: Calling OpenEntry on Session with ulFlags = 0x%X\n", ulFlags);
			hRes = WC_MAPI(lpMAPISession->OpenEntry(cbEntryID, lpEntryID, lpInterface, ulFlags, &ulObjType, &lpUnk));
			if (hRes == MAPI_E_UNKNOWN_FLAGS && ulNoCacheFlags)
			{
				output::DebugPrint(
					DBGGeneric,
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
			auto szFlags = smartview::InterpretNumberAsStringProp(ulObjType, PR_OBJECT_TYPE);
			output::DebugPrint(
				DBGGeneric, L"OnOpenEntryID: Got object of type 0x%08X = %ws\n", ulObjType, szFlags.c_str());
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
		_In_opt_ const SBinary* lpSBinary,
		_In_opt_ LPCIID lpInterface,
		ULONG ulFlags,
		_Out_opt_ ULONG* ulObjTypeRet,
		_Deref_out_opt_ LPUNKNOWN* lppUnk)
	{
		return WC_H(CallOpenEntry(
			lpMDB,
			lpAB,
			lpContainer,
			lpMAPISession,
			lpSBinary ? lpSBinary->cb : 0,
			reinterpret_cast<LPENTRYID>(lpSBinary ? lpSBinary->lpb : nullptr),
			lpInterface,
			ulFlags,
			ulObjTypeRet,
			lppUnk));
	}

	// Concatenate two property arrays without duplicates
	_Check_return_ HRESULT ConcatSPropTagArrays(
		_In_ LPSPropTagArray lpArray1,
		_In_opt_ LPSPropTagArray lpArray2,
		_Deref_out_opt_ LPSPropTagArray* lpNewArray)
	{
		if (!lpNewArray) return MAPI_E_INVALID_PARAMETER;
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
		auto hRes =
			EC_H(MAPIAllocateBuffer(CbNewSPropTagArray(iNewArraySize), reinterpret_cast<LPVOID*>(&lpLocalArray)));

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
			hRes = EC_H(iTargetArray <= iNewArraySize ? S_OK : MAPI_E_CALL_FAILED);

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
	_Check_return_ HRESULT CopyFolderContents(
		_In_ LPMAPIFOLDER lpSrcFolder,
		_In_ LPMAPIFOLDER lpDestFolder,
		bool bCopyAssociatedContents,
		bool bMove,
		bool bSingleCall,
		_In_ HWND hWnd)
	{
		LPMAPITABLE lpSrcContents = nullptr;
		LPSRowSet pRows = nullptr;

		enum
		{
			fldPR_ENTRYID,
			fldNUM_COLS
		};

		static const SizedSPropTagArray(fldNUM_COLS, fldCols) = {
			fldNUM_COLS,
			{PR_ENTRYID},
		};

		output::DebugPrint(
			DBGGeneric,
			L"CopyFolderContents: lpSrcFolder = %p, lpDestFolder = %p, bCopyAssociatedContents = %d, bMove = %d\n",
			lpSrcFolder,
			lpDestFolder,
			bCopyAssociatedContents,
			bMove);

		if (!lpSrcFolder || !lpDestFolder) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(lpSrcFolder->GetContentsTable(
			fMapiUnicode | (bCopyAssociatedContents ? MAPI_ASSOCIATED : NULL), &lpSrcContents));

		if (lpSrcContents)
		{
			hRes = EC_MAPI(lpSrcContents->SetColumns(LPSPropTagArray(&fldCols), TBL_BATCH));

			ULONG ulRowCount = 0;
			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpSrcContents->GetRowCount(0, &ulRowCount));
			}

			if (bSingleCall && ulRowCount < ULONG_MAX / sizeof(SBinary))
			{
				SBinaryArray sbaEID = {0};
				sbaEID.cValues = ulRowCount;
				hRes = EC_H(MAPIAllocateBuffer(sizeof(SBinary) * ulRowCount, reinterpret_cast<LPVOID*>(&sbaEID.lpbin)));
				ZeroMemory(sbaEID.lpbin, sizeof(SBinary) * ulRowCount);

				if (SUCCEEDED(hRes))
				{
					for (ULONG ulRowsCopied = 0; ulRowsCopied < ulRowCount; ulRowsCopied++)
					{
						if (pRows) FreeProws(pRows);
						pRows = nullptr;
						hRes = EC_MAPI(lpSrcContents->QueryRows(1, NULL, &pRows));
						if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

						if (pRows && PT_ERROR != PROP_TYPE(pRows->aRow->lpProps[fldPR_ENTRYID].ulPropTag))
						{
							hRes = EC_H(CopySBinary(
								&sbaEID.lpbin[ulRowsCopied],
								&pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin,
								sbaEID.lpbin));
						}
					}
				}

				LPMAPIPROGRESS lpProgress = mapiui::GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK

				auto ulCopyFlags = bMove ? MESSAGE_MOVE : 0;

				if (lpProgress) ulCopyFlags |= MESSAGE_DIALOG;

				hRes = EC_MAPI(lpSrcFolder->CopyMessages(
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
				if (SUCCEEDED(hRes))
				{
					for (ULONG ulRowsCopied = 0; ulRowsCopied < ulRowCount; ulRowsCopied++)
					{
						if (pRows) FreeProws(pRows);
						pRows = nullptr;
						hRes = EC_MAPI(lpSrcContents->QueryRows(1, NULL, &pRows));
						if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

						if (pRows && PT_ERROR != PROP_TYPE(pRows->aRow->lpProps[fldPR_ENTRYID].ulPropTag))
						{
							SBinaryArray sbaEID = {0};
							output::DebugPrint(DBGGeneric, L"Source Message =\n");
							output::DebugPrintBinary(DBGGeneric, pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin);

							sbaEID.cValues = 1;
							sbaEID.lpbin = &pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin;

							LPMAPIPROGRESS lpProgress =
								mapiui::GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK

							auto ulCopyFlags = bMove ? MESSAGE_MOVE : 0;

							if (lpProgress) ulCopyFlags |= MESSAGE_DIALOG;

							hRes = EC_MAPI(lpSrcFolder->CopyMessages(
								&sbaEID,
								&IID_IMAPIFolder,
								lpDestFolder,
								lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
								lpProgress,
								ulCopyFlags));

							if (hRes == S_OK) output::DebugPrint(DBGGeneric, L"Message Copied\n");

							if (lpProgress) lpProgress->Release();
						}

						if (S_OK != hRes) output::DebugPrint(DBGGeneric, L"Message Copy Failed\n");
					}
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
										if (!lpRows->aRow[iArrayPos].lpProps[ulSrc].Value.bin.cb ||
											!lpRows->aRow[iArrayPos].lpProps[ulSrc].Value.bin.lpb)
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
		LARGE_INTEGER li;
		ULARGE_INTEGER uli;
		ULARGE_INTEGER ulBytesRead;
		ULARGE_INTEGER ulBytesWritten;

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
	_Check_return_ HRESULT CopySBinary(_Out_ LPSBinary psbDest, _In_ const _SBinary* psbSrc, _In_ LPVOID lpParent)
	{
		auto hRes = S_OK;

		if (!psbDest || !psbSrc) return MAPI_E_INVALID_PARAMETER;

		psbDest->cb = psbSrc->cb;

		if (psbSrc->cb)
		{
			psbDest->lpb = mapi::allocate<LPBYTE>(psbSrc->cb, lpParent);

			if (psbDest->lpb) CopyMemory(psbDest->lpb, psbSrc->lpb, psbSrc->cb);
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
	_Check_return_ HRESULT
	CopyStringA(_Deref_out_z_ LPSTR* lpszDestination, _In_z_ LPCSTR szSource, _In_opt_ LPVOID pParent)
	{
		size_t cbSource = 0;

		if (!szSource)
		{
			*lpszDestination = nullptr;
			return S_OK;
		}

		auto hRes = EC_H(StringCbLengthA(szSource, STRSAFE_MAX_CCH * sizeof(char), &cbSource));
		cbSource += sizeof(char);

		if (SUCCEEDED(hRes))
		{
			*lpszDestination = mapi::allocate<LPSTR>(static_cast<ULONG>(cbSource), pParent);
			hRes = EC_H(StringCbCopyA(*lpszDestination, cbSource, szSource));
		}

		return hRes;
	}

	_Check_return_ HRESULT
	CopyStringW(_Deref_out_z_ LPWSTR* lpszDestination, _In_z_ LPCWSTR szSource, _In_opt_ LPVOID pParent)
	{
		size_t cbSource = 0;

		if (!szSource)
		{
			*lpszDestination = nullptr;
			return S_OK;
		}

		auto hRes = EC_H(StringCbLengthW(szSource, STRSAFE_MAX_CCH * sizeof(WCHAR), &cbSource));
		cbSource += sizeof(WCHAR);

		if (SUCCEEDED(hRes))
		{
			*lpszDestination = mapi::allocate<LPWSTR>(static_cast<ULONG>(cbSource), pParent);
		}

		if (SUCCEEDED(hRes))
		{
			hRes = EC_H(StringCbCopyW(*lpszDestination, cbSource, szSource));
		}

		return hRes;
	}

	// Allocates and creates a restriction that looks for existence of
	// a particular property that matches the given string
	// If lpParent is passed in, it is used as the allocation parent.
	_Check_return_ HRESULT CreatePropertyStringRestriction(
		ULONG ulPropTag,
		_In_ const std::wstring& szString,
		ULONG ulFuzzyLevel,
		_In_opt_ LPVOID lpParent,
		_Deref_out_opt_ LPSRestriction* lppRes)
	{
		if (PROP_TYPE(ulPropTag) != PT_UNICODE) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;

		*lppRes = nullptr;

		if (szString.empty()) return MAPI_E_INVALID_PARAMETER;

		// Allocate and create our SRestriction
		// Allocate base memory:
		auto lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), lpParent);
		auto lpAllocationParent = lpParent ? lpParent : lpRes;
		auto lpResLevel1 = mapi::allocate<LPSRestriction>(sizeof(SRestriction) * 2, lpAllocationParent);
		auto lpspvSubject = mapi::allocate<LPSPropValue>(sizeof(SPropValue), lpAllocationParent);

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

			hRes = EC_H(CopyStringW(&lpspvSubject->Value.lpszW, szString.c_str(), lpAllocationParent));

			output::DebugPrint(DBGGeneric, L"CreatePropertyStringRestriction built restriction:\n");
			output::DebugPrintRestriction(DBGGeneric, lpRes, nullptr);

			*lppRes = lpRes;
		}

		if (S_OK != hRes)
		{
			output::DebugPrint(DBGGeneric, L"Failed to create restriction\n");
			MAPIFreeBuffer(lpRes);
			*lppRes = nullptr;
		}

		return hRes;
	}

	_Check_return_ HRESULT CreateRangeRestriction(
		ULONG ulPropTag,
		_In_ const std::wstring& szString,
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
		lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), lpParent);
		lpAllocationParent = lpParent ? lpParent : lpRes;
		lpResLevel1 = mapi::allocate<LPSRestriction>(sizeof(SRestriction) * 2, lpAllocationParent);
		lpspvSubject = mapi::allocate<LPSPropValue>(sizeof(SPropValue), lpAllocationParent);

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

			hRes = EC_H(CopyStringW(&lpspvSubject->Value.lpszW, szString.c_str(), lpAllocationParent));

			output::DebugPrint(DBGGeneric, L"CreateRangeRestriction built restriction:\n");
			output::DebugPrintRestriction(DBGGeneric, lpRes, nullptr);

			*lppRes = lpRes;
		}

		if (S_OK != hRes)
		{
			output::DebugPrint(DBGGeneric, L"Failed to create restriction\n");
			MAPIFreeBuffer(lpRes);
			*lppRes = nullptr;
		}

		return hRes;
	}

	_Check_return_ HRESULT DeleteProperty(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
	{
		LPSPropProblemArray pProbArray = nullptr;

		if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

		if (PROP_TYPE(ulPropTag) == PT_ERROR) ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED);

		output::DebugPrint(
			DBGGeneric, L"DeleteProperty: Deleting prop 0x%08X from MAPI item %p.\n", ulPropTag, lpMAPIProp);

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
		LPMAPIFOLDER lpWasteFolder = nullptr;

		if (!lpMDB || !lpSourceFolder || !lpEIDs) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(
			DBGGeneric, L"DeleteToDeletedItems: Deleting from folder %p in store %p\n", lpSourceFolder, lpMDB);

		auto hRes = WC_H(OpenDefaultFolder(DEFAULT_DELETEDITEMS, lpMDB, &lpWasteFolder));

		if (lpWasteFolder)
		{
			LPMAPIPROGRESS lpProgress = mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::CopyMessages", hWnd); // STRING_OK

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
	FindPropInPropTagArray(_In_ LPSPropTagArray lpspTagArray, ULONG ulPropToFind, _Out_ ULONG* lpulRowFound)
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
		ULONG ulObjType = 0;
		LPSPropValue lpProp = nullptr;

		if (!lpMAPIProp) return 0; // 0's not a valid Object type

		WC_MAPI_S(HrGetOneProp(lpMAPIProp, PR_OBJECT_TYPE, &lpProp));

		if (lpProp) ulObjType = lpProp->Value.ul;

		MAPIFreeBuffer(lpProp);
		return ulObjType;
	}

	_Check_return_ HRESULT GetInbox(_In_ LPMDB lpMDB, _Out_opt_ ULONG* lpcbeid, _Deref_out_opt_ LPENTRYID* lppeid)
	{
		ULONG cbInboxEID = 0;
		LPENTRYID lpInboxEID = nullptr;

		output::DebugPrint(DBGGeneric, L"GetInbox: getting Inbox\n");

		if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(lpMDB->GetReceiveFolder(
			const_cast<LPTSTR>(_T("IPM.Note")), // STRING_OK this is the class of message we want
			fMapiUnicode, // flags
			&cbInboxEID, // size and...
			&lpInboxEID, // value of entry ID
			nullptr)); // returns a message class if not NULL

		if (cbInboxEID && lpInboxEID)
		{
			hRes = WC_H(MAPIAllocateBuffer(cbInboxEID, reinterpret_cast<LPVOID*>(lppeid)));
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
		ULONG cbInboxEID = 0;
		LPENTRYID lpInboxEID = nullptr;

		output::DebugPrint(DBGGeneric, L"GetInbox: getting Inbox from %p\n", lpMDB);

		*lpInbox = nullptr;

		if (!lpMDB || !lpInbox) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_H(GetInbox(lpMDB, &cbInboxEID, &lpInboxEID));

		if (cbInboxEID && lpInboxEID)
		{
			// Get the Inbox...
			hRes = WC_H(CallOpenEntry(
				lpMDB,
				nullptr,
				nullptr,
				nullptr,
				cbInboxEID,
				lpInboxEID,
				nullptr,
				MAPI_BEST_ACCESS,
				nullptr,
				reinterpret_cast<LPUNKNOWN*>(lpInbox)));
		}

		MAPIFreeBuffer(lpInboxEID);
		return hRes;
	}

	_Check_return_ HRESULT
	GetParentFolder(_In_ LPMAPIFOLDER lpChildFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER* lpParentFolder)
	{
		ULONG cProps;
		LPSPropValue lpProps = nullptr;

		*lpParentFolder = nullptr;

		if (!lpChildFolder) return MAPI_E_INVALID_PARAMETER;

		enum
		{
			PARENTEID,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptaSrcFolder) = {NUM_COLS, {PR_PARENT_ENTRYID}};

		// Get PR_PARENT_ENTRYID
		auto hRes =
			EC_H_GETPROPS(lpChildFolder->GetProps(LPSPropTagArray(&sptaSrcFolder), fMapiUnicode, &cProps, &lpProps));

		if (lpProps && PT_ERROR != PROP_TYPE(lpProps[PARENTEID].ulPropTag))
		{
			hRes = WC_H(CallOpenEntry(
				lpMDB,
				nullptr,
				nullptr,
				nullptr,
				lpProps[PARENTEID].Value.bin.cb,
				reinterpret_cast<LPENTRYID>(lpProps[PARENTEID].Value.bin.lpb),
				nullptr,
				MAPI_BEST_ACCESS,
				nullptr,
				reinterpret_cast<LPUNKNOWN*>(lpParentFolder)));
		}

		MAPIFreeBuffer(lpProps);
		return hRes;
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
		if (registry::RegKeys[registry::regkeyUSE_GETPROPLIST].ulCurDWORD)
		{
			output::DebugPrint(DBGGeneric, L"GetPropsNULL: Calling GetPropList\n");
			hRes = WC_MAPI(lpMAPIProp->GetPropList(ulFlags, &lpTags));

			if (hRes == MAPI_E_BAD_CHARWIDTH)
			{
				hRes = EC_MAPI(lpMAPIProp->GetPropList(NULL, &lpTags));
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

	_Check_return_ HRESULT GetSpecialFolderEID(
		_In_ LPMDB lpMDB,
		ULONG ulFolderPropTag,
		_Out_opt_ ULONG* lpcbeid,
		_Deref_out_opt_ LPENTRYID* lppeid)
	{
		if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(DBGGeneric, L"GetSpecialFolderEID: getting 0x%X from %p\n", ulFolderPropTag, lpMDB);

		LPSPropValue lpProp = nullptr;
		LPMAPIFOLDER lpInbox = nullptr;
		auto hRes = WC_H(GetInbox(lpMDB, &lpInbox));
		if (lpInbox)
		{
			hRes = WC_H_MSG(IDS_GETSPECIALFOLDERINBOXMISSINGPROP, HrGetOneProp(lpInbox, ulFolderPropTag, &lpProp));
			lpInbox->Release();
		}

		if (!lpProp)
		{
			LPMAPIFOLDER lpRootFolder = nullptr;
			// Open root container.
			hRes = EC_H(CallOpenEntry(
				lpMDB,
				nullptr,
				nullptr,
				nullptr,
				nullptr, // open root container
				nullptr,
				MAPI_BEST_ACCESS,
				nullptr,
				reinterpret_cast<LPUNKNOWN*>(&lpRootFolder)));
			if (lpRootFolder)
			{
				hRes =
					EC_H_MSG(IDS_GETSPECIALFOLDERROOTMISSINGPROP, HrGetOneProp(lpRootFolder, ulFolderPropTag, &lpProp));
				lpRootFolder->Release();
			}
		}

		if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag) && lpProp->Value.bin.cb)
		{
			hRes = WC_H(MAPIAllocateBuffer(lpProp->Value.bin.cb, reinterpret_cast<LPVOID*>(lppeid)));
			if (SUCCEEDED(hRes))
			{
				*lpcbeid = lpProp->Value.bin.cb;
				CopyMemory(*lppeid, lpProp->Value.bin.lpb, *lpcbeid);
			}
		}

		if (hRes == MAPI_E_NOT_FOUND)
		{
			output::DebugPrint(DBGGeneric, L"Special folder not found.\n");
		}

		MAPIFreeBuffer(lpProp);
		return hRes;
	}

	_Check_return_ HRESULT
	IsAttachmentBlocked(_In_ LPMAPISESSION lpMAPISession, _In_z_ LPCWSTR pwszFileName, _Out_ bool* pfBlocked)
	{
		if (!lpMAPISession || !pwszFileName || !pfBlocked) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		auto bBlocked = BOOL(false);
		auto lpAttachSec = mapi::safe_cast<IAttachmentSecurity*>(lpMAPISession);
		if (lpAttachSec)
		{
			hRes = EC_MAPI(lpAttachSec->IsAttachmentBlocked(pwszFileName, &bBlocked));
			lpAttachSec->Release();
		}

		*pfBlocked = !!bBlocked;
		return hRes;
	}

	_Check_return_ bool IsDuplicateProp(_In_ LPSPropTagArray lpArray, ULONG ulPropTag)
	{
		if (!lpArray) return false;

		for (ULONG i = 0; i < lpArray->cValues; i++)
		{
			// They're dupes if the IDs are the same
			if (registry::RegKeys[registry::regkeyALLOW_DUPE_COLUMNS].ulCurDWORD)
			{
				if (lpArray->aulPropTag[i] == ulPropTag) return true;
			}
			else
			{
				if (PROP_ID(lpArray->aulPropTag[i]) == PROP_ID(ulPropTag)) return true;
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
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

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
							hRes = S_OK;
							ENTRYLIST eid = {0};
							eid.cValues = 1;
							eid.lpbin = &pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID].Value.bin;
							hRes = WC_MAPI(
								lpFolder->DeleteMessages(&eid, NULL, nullptr, bHardDelete ? DELETE_HARD_DELETE : NULL));
							if (SUCCEEDED(hRes)) iItemCount++;
						}
					}
				}
			}

			output::DebugPrint(DBGGeneric, L"ManuallyEmptyFolder deleted %u items\n", iItemCount);
		}

		if (pRows) FreeProws(pRows);
		if (lpContentsTable) lpContentsTable->Release();
		return hRes;
	}

	// Converts vector<BYTE> to LPBYTE allocated with MAPIAllocateMore
	// Will only return nullptr on allocation failure. Even empty bin will return pointer to 0 so MAPI handles empty strings properly
	_Check_return_ LPBYTE ByteVectorToMAPI(const std::vector<BYTE>& bin, LPVOID lpParent)
	{
		// We allocate a couple extra bytes (initialized to NULL) in case this buffer is printed.
		auto lpBin = mapi::allocate<LPBYTE>(static_cast<ULONG>(bin.size()) + sizeof(WCHAR), lpParent);
		if (lpBin)
		{
			memcpy(lpBin, &bin[0], bin.size());
			return lpBin;
		}

		return nullptr;
	}

	ULONG aulOneOffIDs[] = {dispidFormStorage,
							dispidPageDirStream,
							dispidFormPropStream,
							dispidScriptStream,
							dispidPropDefStream, // dispidPropDefStream must remain next to last in list
							dispidCustomFlag}; // dispidCustomFlag must remain last in list

#define ulNumOneOffIDs (_countof(aulOneOffIDs))

	_Check_return_ HRESULT RemoveOneOff(_In_ LPMESSAGE lpMessage, bool bRemovePropDef)
	{
		if (!lpMessage) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(DBGNamedProp, L"RemoveOneOff - removing one off named properties.\n");

		MAPINAMEID rgnmid[ulNumOneOffIDs];
		LPMAPINAMEID rgpnmid[ulNumOneOffIDs];
		LPSPropTagArray lpTags = nullptr;

		for (ULONG i = 0; i < ulNumOneOffIDs; i++)
		{
			rgnmid[i].lpguid = const_cast<LPGUID>(&guid::PSETID_Common);
			rgnmid[i].ulKind = MNID_ID;
			rgnmid[i].Kind.lID = aulOneOffIDs[i];
			rgpnmid[i] = &rgnmid[i];
		}

		auto hRes = EC_H(cache::GetIDsFromNames(lpMessage, ulNumOneOffIDs, rgpnmid, 0, &lpTags));
		if (lpTags)
		{
			LPSPropProblemArray lpProbArray = nullptr;

			output::DebugPrint(DBGNamedProp, L"RemoveOneOff - identified the following properties.\n");
			output::DebugPrintPropTagArray(DBGNamedProp, lpTags);

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
						DBGNamedProp,
						L"RemoveOneOff - DeleteProps problem array:\n%ws\n",
						interpretprop::ProblemArrayToString(*lpProbArray).c_str());
				}

				SPropTagArray pTag = {0};
				ULONG cProp = 0;
				LPSPropValue lpCustomFlag = nullptr;

				// Grab dispidCustomFlag, the last tag in the array
				pTag.cValues = 1;
				pTag.aulPropTag[0] = CHANGE_PROP_TYPE(lpTags->aulPropTag[ulNumOneOffIDs - 1], PT_LONG);

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
							DBGNamedProp,
							L"RemoveOneOff - SetProps problem array:\n%ws\n",
							interpretprop::ProblemArrayToString(*lpProbArray2).c_str());
					}

					MAPIFreeBuffer(lpProbArray2);
				}

				hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
				if (SUCCEEDED(hRes))
				{
					output::DebugPrint(DBGNamedProp, L"RemoveOneOff - One-off properties removed.\n");
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
						LPMESSAGE lpMessage = nullptr;

						hRes = WC_H(CallOpenEntry(
							nullptr,
							nullptr,
							lpFolder,
							nullptr,
							pRows->aRow[i].lpProps[ePR_ENTRYID].Value.bin.cb,
							reinterpret_cast<LPENTRYID>(pRows->aRow[i].lpProps[ePR_ENTRYID].Value.bin.lpb),
							nullptr,
							MAPI_BEST_ACCESS,
							nullptr,
							reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
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
		LPMESSAGE lpMessage = nullptr;

		if (!lpFolder || !MessageEID) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_H(CallOpenEntry(
			nullptr,
			nullptr,
			lpFolder,
			nullptr,
			MessageEID->cb,
			reinterpret_cast<LPENTRYID>(MessageEID->lpb),
			nullptr,
			MAPI_BEST_ACCESS,
			nullptr,
			reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
		if (lpMessage)
		{
			hRes = EC_H(ResendSingleMessage(lpFolder, lpMessage, hWnd));
		}

		if (lpMessage) lpMessage->Release();
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
		SPropValue sProp = {0};

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

		output::DebugPrint(DBGGeneric, L"ResendSingleMessage: Checking message for embedded messages\n");

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
					if (pRows && (!pRows || pRows->cRows)) break;

					if (ATTACH_EMBEDDED_MSG == pRows->aRow->lpProps[atPR_ATTACH_METHOD].Value.l)
					{
						output::DebugPrint(DBGGeneric, L"Found an embedded message to resend.\n");

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

						output::DebugPrint(DBGGeneric, L"Message opened.\n");

						if (CheckStringProp(&pRows->aRow->lpProps[atPR_DISPLAY_NAME], PT_UNICODE))
						{
							output::DebugPrint(
								DBGGeneric,
								L"Resending \"%ws\"\n",
								pRows->aRow->lpProps[atPR_DISPLAY_NAME].Value.lpszW);
						}

						output::DebugPrint(DBGGeneric, L"Creating new message.\n");
						if (lpNewMessage) lpNewMessage->Release();
						lpNewMessage = nullptr;
						hRes = EC_MAPI(lpFolder->CreateMessage(nullptr, 0, &lpNewMessage));
						if (FAILED(hRes) || !lpNewMessage) continue;

						hRes = EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));
						if (FAILED(hRes)) continue;

						// Copy all the transmission properties
						output::DebugPrint(DBGGeneric, L"Getting list of properties.\n");
						MAPIFreeBuffer(lpsMessageTags);
						lpsMessageTags = nullptr;
						hRes = EC_MAPI(lpAttachMsg->GetPropList(0, &lpsMessageTags));
						if (FAILED(hRes) || !!lpsMessageTags) continue;

						output::DebugPrint(DBGGeneric, L"Copying properties to new message.\n");
						if (SUCCEEDED(hRes))
						{
							for (ULONG ulProp = 0; ulProp < lpsMessageTags->cValues; ulProp++)
							{
								// it would probably be quicker to use this loop to construct an array of properties
								// we desire to copy, and then pass that array to GetProps and then SetProps
								if (FIsTransmittable(lpsMessageTags->aulPropTag[ulProp]))
								{
									LPSPropValue lpProp = nullptr;
									output::DebugPrint(
										DBGGeneric, L"Copying 0x%08X\n", lpsMessageTags->aulPropTag[ulProp]);
									hRes = WC_MAPI(HrGetOnePropEx(
										lpAttachMsg, lpsMessageTags->aulPropTag[ulProp], fMapiUnicode, &lpProp));

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

						output::DebugPrint(DBGGeneric, L"Copying recipients and attachments to new message.\n");

						LPMAPIPROGRESS lpProgress =
							mapi::mapiui::GetMAPIProgress(L"IMAPIProp::CopyProps", hWnd); // STRING_OK

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

						output::DebugPrint(DBGGeneric, L"Setting PR_DELETE_AFTER_SUBMIT to true.\n");
						hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));
						if (FAILED(hRes)) continue;

						SPropTagArray sPropTagArray = {0};

						sPropTagArray.cValues = 1;
						sPropTagArray.aulPropTag[0] = PR_SENTMAIL_ENTRYID;

						output::DebugPrint(DBGGeneric, L"Deleting PR_SENTMAIL_ENTRYID\n");
						hRes = EC_MAPI(lpNewMessage->DeleteProps(&sPropTagArray, nullptr));
						if (FAILED(hRes)) continue;

						hRes = EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));
						if (FAILED(hRes)) continue;

						output::DebugPrint(DBGGeneric, L"Submitting new message.\n");
						hRes = EC_MAPI(lpNewMessage->SubmitMessage(0));
						if (FAILED(hRes)) continue;
					}
					else
					{
						output::DebugPrint(DBGGeneric, L"Attachment is not an embedded message.\n");
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
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

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
							lpMessage = nullptr;

							hRes = WC_H(CallOpenEntry(
								lpMDB,
								nullptr,
								nullptr,
								nullptr,
								pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID].Value.bin.cb,
								reinterpret_cast<LPENTRYID>(
									pRows->aRow[iCurPropRow].lpProps[eidPR_ENTRYID].Value.bin.lpb),
								nullptr,
								MAPI_BEST_ACCESS,
								nullptr,
								reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
							if (FAILED(hRes))
							{
								hResOverall = hRes;
								continue;
							}

							hRes = WC_H(DeleteProperty(lpMessage, PR_NT_SECURITY_DESCRIPTOR));
							if (FAILED(hRes))
							{
								hResOverall = hRes;
								continue;
							}

							iItemCount++;
						}
					}
				}

				output::DebugPrint(DBGGeneric, L"ResetPermissionsOnItems reset permissions on %u items\n", iItemCount);
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
			SPropValue sProp;

			sProp.dwAlignPad = 0;
			sProp.ulPropTag = PR_DELETE_AFTER_SUBMIT;
			sProp.Value.b = true;

			output::DebugPrint(DBGGeneric, L"Setting PR_DELETE_AFTER_SUBMIT to true.\n");
			hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));

			if (SUCCEEDED(hRes))
			{
				sProp.dwAlignPad = 0;
				sProp.ulPropTag = PR_BODY_W;
				sProp.Value.lpszW = const_cast<LPWSTR>(szBody.c_str());

				output::DebugPrint(DBGGeneric, L"Setting PR_BODY to %ws.\n", szBody.c_str());
				hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));
			}

			if (SUCCEEDED(hRes))
			{
				sProp.dwAlignPad = 0;
				sProp.ulPropTag = PR_SUBJECT_W;
				sProp.Value.lpszW = const_cast<LPWSTR>(szSubject.c_str());

				output::DebugPrint(DBGGeneric, L"Setting PR_SUBJECT to %ws.\n", szSubject.c_str());
				hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));
			}

			if (SUCCEEDED(hRes))
			{
				sProp.dwAlignPad = 0;
				sProp.ulPropTag = PR_MESSAGE_CLASS_W;
				sProp.Value.lpszW = const_cast<LPWSTR>(szClass.c_str());

				output::DebugPrint(DBGGeneric, L"Setting PR_MESSAGE_CLASS to %ws.\n", szSubject.c_str());
				hRes = EC_MAPI(HrSetOneProp(lpNewMessage, &sProp));
			}

			if (SUCCEEDED(hRes))
			{
				SPropTagArray sPropTagArray;

				sPropTagArray.cValues = 1;
				sPropTagArray.aulPropTag[0] = PR_SENTMAIL_ENTRYID;

				output::DebugPrint(DBGGeneric, L"Deleting PR_SENTMAIL_ENTRYID\n");
				hRes = EC_MAPI(lpNewMessage->DeleteProps(&sPropTagArray, nullptr));
			}

			if (SUCCEEDED(hRes))
			{
				output::DebugPrint(DBGGeneric, L"Adding recipient: %ws.\n", szRecipient.c_str());
				hRes = EC_H(ab::AddRecipient(lpMAPISession, lpNewMessage, szRecipient, MAPI_TO));
			}

			if (SUCCEEDED(hRes))
			{
				output::DebugPrint(DBGGeneric, L"Submitting message\n");
				hRes = EC_MAPI(lpNewMessage->SubmitMessage(NULL));
			}
		}

		if (lpNewMessage) lpNewMessage->Release();
		return hRes;
	}

	// Declaration missing from headers
	STDAPI_(HRESULT)
	WrapCompressedRTFStreamEx(
		LPSTREAM pCompressedRTFStream,
		const RTF_WCSINFO* pWCSInfo,
		LPSTREAM* ppUncompressedRTFStream,
		RTF_WCSRETINFO* pRetInfo);

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
			hRes = WC_MAPI(WrapCompressedRTFStream(lpCompressedRTFStream, ulFlags, lpUncompressedRTFStream));
		}
		else
		{
			RTF_WCSINFO wcsinfo = {0};
			RTF_WCSRETINFO retinfo = {0};

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

		LPSPropTagArray lpPropTags = nullptr;

		auto hRes = EC_H(GetNamedPropsByGUID(lpSource, lpPropSetGUID, &lpPropTags));

		if (SUCCEEDED(hRes) && lpPropTags)
		{
			LPSPropProblemArray lpProblems = nullptr;
			ULONG ulFlags = 0;
			if (bDoMove) ulFlags |= MAPI_MOVE;
			if (bDoNoReplace) ulFlags |= MAPI_NOREPLACE;

			LPMAPIPROGRESS lpProgress = mapi::mapiui::GetMAPIProgress(L"IMAPIProp::CopyProps", hWnd); // STRING_OK

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

	_Check_return_ HRESULT
	GetNamedPropsByGUID(_In_ LPMAPIPROP lpSource, _In_ LPGUID lpPropSetGUID, _Deref_out_ LPSPropTagArray* lpOutArray)
	{
		if (!lpSource || !lpPropSetGUID || lpOutArray) return MAPI_E_INVALID_PARAMETER;

		LPSPropTagArray lpAllProps = nullptr;

		*lpOutArray = nullptr;

		auto hRes = WC_MAPI(lpSource->GetPropList(0, &lpAllProps));

		if (hRes == S_OK && lpAllProps)
		{
			ULONG cProps = 0;
			LPMAPINAMEID* lppNameIDs = nullptr;

			hRes = WC_H(cache::GetNamesFromIDs(lpSource, &lpAllProps, nullptr, 0, &cProps, &lppNameIDs));

			if (hRes == S_OK && lppNameIDs)
			{
				ULONG ulNumProps = 0; // count of props that match our guid
				for (ULONG i = 0; i < cProps; i++)
				{
					if (PROP_ID(lpAllProps->aulPropTag[i]) > 0x7FFF && lppNameIDs[i] &&
						!memcmp(lppNameIDs[i]->lpguid, lpPropSetGUID, sizeof(GUID)))
					{
						ulNumProps++;
					}
				}

				LPSPropTagArray lpFilteredProps = nullptr;

				hRes = WC_H(
					MAPIAllocateBuffer(CbNewSPropTagArray(ulNumProps), reinterpret_cast<LPVOID*>(&lpFilteredProps)));

				if (hRes == S_OK && lpFilteredProps)
				{
					lpFilteredProps->cValues = 0;

					for (ULONG i = 0; i < cProps; i++)
					{
						if (PROP_ID(lpAllProps->aulPropTag[i]) > 0x7FFF && lppNameIDs[i] &&
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

	_Check_return_ bool CheckStringProp(_In_opt_ const _SPropValue* lpProp, ULONG ulPropType)
	{
		if (PT_STRING8 != ulPropType && PT_UNICODE != ulPropType)
		{
			output::DebugPrint(DBGGeneric, L"CheckStringProp: Called with invalid ulPropType of 0x%X\n", ulPropType);
			return false;
		}

		if (!lpProp)
		{
			output::DebugPrint(DBGGeneric, L"CheckStringProp: lpProp is NULL\n");
			return false;
		}

		if (PT_ERROR == PROP_TYPE(lpProp->ulPropTag))
		{
			output::DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->ulPropTag is of type PT_ERROR\n");
			return false;
		}

		if (ulPropType != PROP_TYPE(lpProp->ulPropTag))
		{
			output::DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->ulPropTag is not of type 0x%X\n", ulPropType);
			return false;
		}

		if (nullptr == lpProp->Value.LPSZ)
		{
			output::DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->Value.LPSZ is NULL\n");
			return false;
		}

		if (PT_STRING8 == ulPropType && NULL == lpProp->Value.lpszA[0])
		{
			output::DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->Value.lpszA[0] is NULL\n");
			return false;
		}

		if (PT_UNICODE == ulPropType && NULL == lpProp->Value.lpszW[0])
		{
			output::DebugPrint(DBGGeneric, L"CheckStringProp: lpProp->Value.lpszW[0] is NULL\n");
			return false;
		}

		return true;
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
			output::DebugPrint(DBGGeneric, L"ComputeStoreHash, hash (before adding .PUB) = 0x%08X\n", dwHash);
			// augment to make sure it is unique else could be same as the private store
			dwHash = (dwHash << 5) + dwHash + 0x2E505542; // this is '.PUB'
		}

		if (pwzFileName || pszFileName)
			output::DebugPrint(DBGGeneric, L"ComputeStoreHash, hash (before adding path) = 0x%08X\n", dwHash);

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
			output::DebugPrint(DBGGeneric, L"ComputeStoreHash, hash (after adding path) = 0x%08X\n", dwHash);

		// dwHash now contains the hash to be used. It should be written in hex when building a URL.
		return dwHash;
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

	HRESULT
	HrEmsmdbUIDFromStore(_In_ LPMAPISESSION pmsess, _In_ const MAPIUID* puidService, _Out_opt_ MAPIUID* pEmsmdbUID)
	{
		if (!puidService) return MAPI_E_INVALID_PARAMETER;

		SRestriction mres = {0};
		SPropValue mval = {0};
		SRowSet* pRows = nullptr;
		LPSERVICEADMIN spSvcAdmin = nullptr;
		LPMAPITABLE spmtab = nullptr;

		enum
		{
			eEntryID = 0,
			eSectionUid,
			eMax
		};
		static const SizedSPropTagArray(eMax, tagaCols) = {eMax,
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
					mval.Value.bin.cb = sizeof *puidService;
					mval.Value.bin.lpb = LPBYTE(puidService);

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
						if (PR_EMSMDB_SECTION_UID == pRow->lpProps[eSectionUid].ulPropTag &&
							pRow->lpProps[eSectionUid].Value.bin.cb == sizeof *pEmsmdbUID)
						{
							memcpy(pEmsmdbUID, pRow->lpProps[eSectionUid].Value.bin.lpb, sizeof *pEmsmdbUID);
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
		return IsEqualMAPIUID(lpmapiuid, LPMAPIUID(pbExchangeProviderPrimaryUserGuid));
	}

	bool FExchangePublicStore(_In_ LPMAPIUID lpmapiuid)
	{
		if (!lpmapiuid) return false;
		return IsEqualMAPIUID(lpmapiuid, LPMAPIUID(pbExchangeProviderPublicGuid));
	}

	STDMETHODIMP
	GetEntryIDFromMDB(LPMDB lpMDB, ULONG ulPropTag, _Out_opt_ ULONG* lpcbeid, _Deref_out_opt_ LPENTRYID* lppeid)
	{
		if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
		LPSPropValue lpEIDProp = nullptr;

		auto hRes = WC_MAPI(HrGetOneProp(lpMDB, ulPropTag, &lpEIDProp));

		if (SUCCEEDED(hRes) && lpEIDProp)
		{
			hRes = WC_H(MAPIAllocateBuffer(lpEIDProp->Value.bin.cb, reinterpret_cast<LPVOID*>(lppeid)));
			if (SUCCEEDED(hRes))
			{
				*lpcbeid = lpEIDProp->Value.bin.cb;
				CopyMemory(*lppeid, lpEIDProp->Value.bin.lpb, *lpcbeid);
			}
		}

		MAPIFreeBuffer(lpEIDProp);
		return hRes;
	}

	STDMETHODIMP GetMVEntryIDFromInboxByIndex(
		LPMDB lpMDB,
		ULONG ulPropTag,
		ULONG ulIndex,
		_Out_opt_ ULONG* lpcbeid,
		_Deref_out_opt_ LPENTRYID* lppeid)
	{
		if (!lpMDB || !lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
		LPMAPIFOLDER lpInbox = nullptr;

		auto hRes = WC_H(GetInbox(lpMDB, &lpInbox));

		if (SUCCEEDED(hRes) && lpInbox)
		{
			LPSPropValue lpEIDProp = nullptr;
			hRes = WC_MAPI(HrGetOneProp(lpInbox, ulPropTag, &lpEIDProp));
			if (SUCCEEDED(hRes) && lpEIDProp && PT_MV_BINARY == PROP_TYPE(lpEIDProp->ulPropTag) &&
				ulIndex < lpEIDProp->Value.MVbin.cValues && lpEIDProp->Value.MVbin.lpbin[ulIndex].cb > 0)
			{
				hRes = WC_H(
					MAPIAllocateBuffer(lpEIDProp->Value.MVbin.lpbin[ulIndex].cb, reinterpret_cast<LPVOID*>(lppeid)));
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

	STDMETHODIMP
	GetDefaultFolderEID(
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

	STDMETHODIMP OpenDefaultFolder(_In_ ULONG ulFolder, _In_ LPMDB lpMDB, _Deref_out_opt_ LPMAPIFOLDER* lpFolder)
	{
		if (!lpMDB || !lpFolder) return MAPI_E_INVALID_PARAMETER;

		*lpFolder = nullptr;
		ULONG cb = 0;
		LPENTRYID lpeid = nullptr;

		auto hRes = WC_H(GetDefaultFolderEID(ulFolder, lpMDB, &cb, &lpeid));
		if (SUCCEEDED(hRes))
		{
			LPMAPIFOLDER lpTemp = nullptr;
			hRes = WC_H(CallOpenEntry(
				lpMDB,
				nullptr,
				nullptr,
				nullptr,
				cb,
				lpeid,
				nullptr,
				MAPI_BEST_ACCESS,
				nullptr,
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
			szTitle = strings::loadstring(IDS_DISPLAYNAMENOTFOUND);
		}

		return szTitle;
	}

	bool UnwrapContactEntryID(_In_ ULONG cbIn, _In_ LPBYTE lpbIn, _Out_ ULONG* lpcbOut, _Out_ LPBYTE* lppbOut)
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
		dialog::editor::CTagArrayEditor TagEditor(
			nullptr, IDS_TAGSTOEXCLUDE, IDS_TAGSTOEXCLUDEPROMPT, nullptr, lpTagArray, bIsAB, lpProp);
		if (TagEditor.DisplayDialog())
		{
			return TagEditor.DetachModifiedTagArray();
		}

		return nullptr;
	}
#endif

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
			dialog::editor::CEditor MyData(
				nullptr, IDS_COPYTO, IDS_COPYPASTEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.InitPane(
				0, viewpane::TextPane::CreateSingleLinePane(IDS_INTERFACE, guid::GUIDToStringAndName(lpGUID), false));
			MyData.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_FLAGS, false));
			MyData.SetHex(1, MAPI_DIALOG);

			if (!MyData.DisplayDialog()) return MAPI_E_USER_CANCEL;

			auto MyGUID = guid::StringToGUID(MyData.GetStringW(0));
			lpGUIDLocal = &MyGUID;
			ulCopyFlags = MyData.GetHex(1);
			if (hWnd)
			{
				lpProgress = mapi::mapiui::GetMAPIProgress(L"CopyTo", hWnd); // STRING_OK
				if (lpProgress) ulCopyFlags |= MAPI_DIALOG;
			}

			lpUITags = GetExcludedTags(lpTagArray, lpSource, bIsAB);
			if (lpUITags)
			{
				lpExcludedTags = lpUITags;
			}
		}
#endif

		auto hRes = WC_MAPI(lpSource->CopyTo(
			0,
			nullptr,
			lpExcludedTags,
			lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL, // UI param
			lpProgress, // progress
			lpGUIDLocal,
			lpDest,
			ulCopyFlags, // flags
			&lpProblems));

		MAPIFreeBuffer(lpUITags);
		if (lpProgress) lpProgress->Release();

		if (lpProblems)
		{
			WC_PROBLEMARRAY(lpProblems);
			MAPIFreeBuffer(lpProblems);
		}

		return hRes;
	}

	// Augemented version of HrGetOneProp which allows passing flags to underlying GetProps
	// Useful for passing fMapiUnicode for unspecified string/stream types
	HRESULT
	HrGetOnePropEx(_In_ LPMAPIPROP lpMAPIProp, _In_ ULONG ulPropTag, _In_ ULONG ulFlags, _Out_ LPSPropValue* lppProp)
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

	// Returns LPSPropValue with value of a property
	// Uses GetProps and falls back to OpenProperty if the value is large
	// Free with MAPIFreeBuffer
	_Check_return_ HRESULT
	GetLargeProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp)
	{
		if (!lpMAPIProp || !lppProp) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(DBGGeneric, L"GetLargeProp getting buffer from 0x%08X\n", ulPropTag);

		ULONG cValues = 0;
		LPSPropValue lpPropArray = nullptr;
		auto bSuccess = false;

		const SizedSPropTagArray(1, sptaBuffer) = {1, {ulPropTag}};
		*lppProp = nullptr;

		auto hRes = WC_H_GETPROPS(lpMAPIProp->GetProps(LPSPropTagArray(&sptaBuffer), 0, &cValues, &lpPropArray));

		if (lpPropArray && PT_ERROR == PROP_TYPE(lpPropArray->ulPropTag) &&
			lpPropArray->Value.err == MAPI_E_NOT_ENOUGH_MEMORY)
		{
			output::DebugPrint(DBGGeneric, L"GetLargeProp property reported in GetProps as large.\n");
			MAPIFreeBuffer(lpPropArray);
			lpPropArray = nullptr;
			// need to get the data as a stream
			LPSTREAM lpStream = nullptr;

			hRes = WC_MAPI(lpMAPIProp->OpenProperty(
				ulPropTag, &IID_IStream, STGM_READ, 0, reinterpret_cast<LPUNKNOWN*>(&lpStream)));
			if (SUCCEEDED(hRes) && lpStream)
			{
				STATSTG StatInfo = {nullptr};
				lpStream->Stat(&StatInfo, STATFLAG_NONAME); // find out how much space we need

				// We're not going to try to support MASSIVE properties.
				if (!StatInfo.cbSize.HighPart)
				{
					hRes = EC_H(MAPIAllocateBuffer(sizeof(SPropValue), reinterpret_cast<LPVOID*>(&lpPropArray)));
					if (lpPropArray)
					{
						memset(lpPropArray, 0, sizeof(SPropValue));
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
								hRes = EC_MAPI(lpStream->Read(lpBuffer, ulBufferSize, &ulSizeRead));
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

		if (bSuccess)
		{
			*lppProp = lpPropArray;
		}
		else
		{
			MAPIFreeBuffer(lpPropArray);
			if (SUCCEEDED(hRes)) hRes = MAPI_E_CALL_FAILED;
		}

		return hRes;
	}

	// Returns LPSPropValue with value of a binary property
	// Free with MAPIFreeBuffer
	_Check_return_ HRESULT
	GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp)
	{
		return GetLargeProp(lpMAPIProp, CHANGE_PROP_TYPE(ulPropTag, PT_BINARY), lppProp);
	}

	// Returns LPSPropValue with value of a string property
	// Free with MAPIFreeBuffer
	_Check_return_ HRESULT
	GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp)
	{
		return GetLargeProp(lpMAPIProp, CHANGE_PROP_TYPE(ulPropTag, PT_TSTRING), lppProp);
	}

	_Check_return_ HRESULT
	HrDupPropset(int cprop, _In_count_(cprop) LPSPropValue rgprop, _In_ LPVOID lpObject, _In_ LPSPropValue* prgprop)
	{
		ULONG cb = NULL;

		// Find out how much memory we need
		auto hRes = EC_MAPI(ScCountProps(cprop, rgprop, &cb));

		if (SUCCEEDED(hRes) && cb)
		{
			// Obtain memory
			*prgprop = mapi::allocate<LPSPropValue>(cb, lpObject);
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
		_In_opt_ const LPVOID lpObject, // ptr to existing MAPI buffer
		_In_ LPSRestriction* lppResDest) // dest restriction buffer ptr
	{
		if (!lppResDest) return MAPI_E_INVALID_PARAMETER;
		*lppResDest = nullptr;
		if (!lpResSrc) return S_OK;

		*lppResDest = mapi::allocate<LPSRestriction>(sizeof(SRestriction), lpObject);
		auto lpAllocationParent = lpObject ? lpObject : *lppResDest;
		auto hRes = WC_H(HrCopyRestrictionArray(lpResSrc, lpAllocationParent, 1, *lppResDest));

		if (FAILED(hRes))
		{
			if (!lpObject) MAPIFreeBuffer(*lppResDest);
		}

		return hRes;
	}

	_Check_return_ HRESULT HrCopyRestrictionArray(
		_In_ const _SRestriction* lpResSrc, // source restriction
		_In_ const LPVOID lpObject, // ptr to existing MAPI buffer
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
						1, lpResSrc[i].res.resContent.lpProp, lpObject, &lpResDest[i].res.resContent.lpProp));
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
						lpObject,
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

	typedef ACTIONS* LPACTIONS;

	// swiped from EDK rules sample
	_Check_return_ STDAPI HrCopyActions(
		_In_ LPACTIONS lpActsSrc, // source action ptr
		_In_ LPVOID lpObject, // ptr to existing MAPI buffer
		_In_ LPACTIONS* lppActsDst) // ptr to destination ACTIONS buffer
	{
		if (!lpActsSrc || !lppActsDst) return MAPI_E_INVALID_PARAMETER;
		if (lpActsSrc->cActions <= 0 || lpActsSrc->lpAction == nullptr) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;

		*lppActsDst = mapi::allocate<LPACTIONS>(sizeof(ACTIONS), lpObject);
		auto lpAllocationParent = lpObject ? lpObject : *lppActsDst;
		// no short circuit returns after here

		auto lpActsDst = *lppActsDst;
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
								lpObject,
								&lpActDst->lpadrlist->aEntries[j].rgPropVals));
							if (FAILED(hRes)) break;
						}
					}
					break;

				case OP_TAG: // propTag
					hRes = WC_H(MyPropCopyMore(&lpActDst->propTag, &lpActSrc->propTag, MAPIAllocateMore, lpObject));
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
			if (!lpObject) MAPIFreeBuffer(*lppActsDst);
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
		_In_ LPVOID lpvObject)
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
					reinterpret_cast<LPSRestriction>(lpSPropValueSrc->Value.lpszA), lpvObject, &lpNewRes));
				lpSPropValueDest->Value.lpszA = reinterpret_cast<LPSTR>(lpNewRes);
			}
			else
			{
				ACTIONS* lpNewAct = nullptr;
				hRes =
					WC_H(HrCopyActions(reinterpret_cast<ACTIONS*>(lpSPropValueSrc->Value.lpszA), lpvObject, &lpNewAct));
				lpSPropValueDest->Value.lpszA = reinterpret_cast<LPSTR>(lpNewAct);
			}
			break;
		}
		default:
			hRes = WC_MAPI(
				PropCopyMore(lpSPropValueDest, const_cast<LPSPropValue>(lpSPropValueSrc), lpfAllocMore, lpvObject));
		}

		return hRes;
	}

	// Declaration missing from MAPI headers
	_Check_return_ STDAPI OpenStreamOnFileW(
		_In_ LPALLOCATEBUFFER lpAllocateBuffer,
		_In_ LPFREEBUFFER lpFreeBuffer,
		ULONG ulFlags,
		_In_z_ LPCWSTR lpszFileName,
		_In_opt_z_ LPCWSTR lpszPrefix,
		_Out_ LPSTREAM FAR* lppStream);

	// Since I never use lpszPrefix, I don't convert it
	// To make certain of that, I pass NULL for it
	// If I ever do need this param, I'll have to fix this
	_Check_return_ STDMETHODIMP MyOpenStreamOnFile(
		_In_ LPALLOCATEBUFFER lpAllocateBuffer,
		_In_ LPFREEBUFFER lpFreeBuffer,
		ULONG ulFlags,
		_In_ const std::wstring& lpszFileName,
		_Out_ LPSTREAM FAR* lppStream)
	{
		auto hRes =
			OpenStreamOnFileW(lpAllocateBuffer, lpFreeBuffer, ulFlags, lpszFileName.c_str(), nullptr, lppStream);
		if (hRes == MAPI_E_CALL_FAILED)
		{
			hRes = OpenStreamOnFile(
				lpAllocateBuffer,
				lpFreeBuffer,
				ulFlags,
				strings::wstringTotstring(lpszFileName).c_str(),
				nullptr,
				lppStream);
		}

		return hRes;
	}
}