// Collection of useful MAPI Address Book functions

#include <StdAfx.h>
#include <MAPI/MAPIABFunctions.h>
#include <MAPI/MAPIFunctions.h>

namespace mapi
{
	namespace ab
	{
		_Check_return_ HRESULT HrAllocAdrList(ULONG ulNumProps, _Deref_out_opt_ LPADRLIST* lpAdrList)
		{
			if (!lpAdrList || ulNumProps > ULONG_MAX / sizeof(SPropValue)) return MAPI_E_INVALID_PARAMETER;
			auto hRes = S_OK;
			LPADRLIST lpLocalAdrList = nullptr;

			*lpAdrList = nullptr;

			// Allocate memory for new SRowSet structure.
			EC_H(MAPIAllocateBuffer(CbNewSRowSet(1), reinterpret_cast<LPVOID*>(&lpLocalAdrList)));

			if (lpLocalAdrList)
			{
				// Zero out allocated memory.
				ZeroMemory(lpLocalAdrList, CbNewSRowSet(1));

				// Allocate memory for SPropValue structure that indicates what
				// recipient properties will be set.
				EC_H(MAPIAllocateBuffer(
					ulNumProps * sizeof(SPropValue),
					reinterpret_cast<LPVOID*>(&lpLocalAdrList->aEntries[0].rgPropVals)));

				// Zero out allocated memory.
				if (lpLocalAdrList->aEntries[0].rgPropVals)
					ZeroMemory(lpLocalAdrList->aEntries[0].rgPropVals, ulNumProps * sizeof(SPropValue));
				if (SUCCEEDED(hRes))
				{
					*lpAdrList = lpLocalAdrList;
				}
				else
				{
					FreePadrlist(lpLocalAdrList);
				}
			}

			return hRes;
		}

		_Check_return_ HRESULT AddOneOffAddress(
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMESSAGE lpMessage,
			_In_ const std::wstring& szDisplayName,
			_In_ const std::wstring& szAddrType,
			_In_ const std::wstring& szEmailAddress,
			ULONG ulRecipientType)
		{
			auto hRes = S_OK;
			LPADRLIST lpAdrList = nullptr; // ModifyRecips takes LPADRLIST
			LPADRBOOK lpAddrBook = nullptr;
			LPENTRYID lpEID = nullptr;

			enum
			{
				NAME,
				ADDR,
				EMAIL,
				RECIP,
				EID,
				NUM_RECIP_PROPS
			};

			if (!lpMessage || !lpMAPISession) return MAPI_E_INVALID_PARAMETER;

			EC_MAPI(lpMAPISession->OpenAddressBook(
				NULL,
				nullptr,
				NULL,
				&lpAddrBook));

			EC_MAPI(HrAllocAdrList(NUM_RECIP_PROPS, &lpAdrList));

			// Setup the One Time recipient by indicating how many recipients
			// and how many properties will be set on each recipient.

			if (SUCCEEDED(hRes) && lpAdrList)
			{
				lpAdrList->cEntries = 1; // How many recipients.
				lpAdrList->aEntries[0].cValues = NUM_RECIP_PROPS; // How many properties per recipient

				// Set the SPropValue members == the desired values.
				lpAdrList->aEntries[0].rgPropVals[NAME].ulPropTag = PR_DISPLAY_NAME_W;
				lpAdrList->aEntries[0].rgPropVals[NAME].Value.lpszW = const_cast<LPWSTR>(szDisplayName.c_str());

				lpAdrList->aEntries[0].rgPropVals[ADDR].ulPropTag = PR_ADDRTYPE_W;
				lpAdrList->aEntries[0].rgPropVals[ADDR].Value.lpszW = const_cast<LPWSTR>(szAddrType.c_str());

				lpAdrList->aEntries[0].rgPropVals[EMAIL].ulPropTag = PR_EMAIL_ADDRESS_W;
				lpAdrList->aEntries[0].rgPropVals[EMAIL].Value.lpszW = const_cast<LPWSTR>(szEmailAddress.c_str());

				lpAdrList->aEntries[0].rgPropVals[RECIP].ulPropTag = PR_RECIPIENT_TYPE;
				lpAdrList->aEntries[0].rgPropVals[RECIP].Value.l = ulRecipientType;

				lpAdrList->aEntries[0].rgPropVals[EID].ulPropTag = PR_ENTRYID;

				// Create the One-off address and get an EID for it.
				EC_MAPI(lpAddrBook->CreateOneOff(
					reinterpret_cast<LPTSTR>(lpAdrList->aEntries[0].rgPropVals[NAME].Value.lpszW),
					reinterpret_cast<LPTSTR>(lpAdrList->aEntries[0].rgPropVals[ADDR].Value.lpszW),
					reinterpret_cast<LPTSTR>(lpAdrList->aEntries[0].rgPropVals[EMAIL].Value.lpszW),
					MAPI_UNICODE,
					&lpAdrList->aEntries[0].rgPropVals[EID].Value.bin.cb,
					&lpEID));
				lpAdrList->aEntries[0].rgPropVals[EID].Value.bin.lpb = reinterpret_cast<LPBYTE>(lpEID);

				EC_MAPI(lpAddrBook->ResolveName(
					0L,
					MAPI_UNICODE,
					nullptr,
					lpAdrList));

				// If everything goes right, add the new recipient to the message
				// object passed into us.
				EC_MAPI(lpMessage->ModifyRecipients(MODRECIP_ADD, lpAdrList));

				EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			}

			MAPIFreeBuffer(lpEID);
			if (lpAdrList) FreePadrlist(lpAdrList);
			if (lpAddrBook) lpAddrBook->Release();
			return hRes;
		}

		_Check_return_ HRESULT AddRecipient(
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMESSAGE lpMessage,
			_In_ const std::wstring& szName,
			ULONG ulRecipientType)
		{
			auto hRes = S_OK;
			LPADRLIST lpAdrList = nullptr; // ModifyRecips takes LPADRLIST
			LPADRBOOK lpAddrBook = nullptr;

			enum
			{
				NAME,
				RECIP,
				NUM_RECIP_PROPS
			};

			if (!lpMessage || !lpMAPISession) return MAPI_E_INVALID_PARAMETER;

			EC_MAPI(lpMAPISession->OpenAddressBook(
				NULL,
				nullptr,
				NULL,
				&lpAddrBook));

			EC_MAPI(HrAllocAdrList(NUM_RECIP_PROPS, &lpAdrList));

			if (lpAdrList)
			{
				// Setup the One Time recipient by indicating how many recipients
				// and how many properties will be set on each recipient.
				lpAdrList->cEntries = 1; // How many recipients.
				lpAdrList->aEntries[0].cValues = NUM_RECIP_PROPS; // How many properties per recipient

				// Set the SPropValue members == the desired values.
				lpAdrList->aEntries[0].rgPropVals[NAME].ulPropTag = PR_DISPLAY_NAME_W;
				lpAdrList->aEntries[0].rgPropVals[NAME].Value.lpszW = const_cast<LPWSTR>(szName.c_str());

				lpAdrList->aEntries[0].rgPropVals[RECIP].ulPropTag = PR_RECIPIENT_TYPE;
				lpAdrList->aEntries[0].rgPropVals[RECIP].Value.l = ulRecipientType;

				EC_MAPI(lpAddrBook->ResolveName(
					0L,
					fMapiUnicode,
					nullptr,
					lpAdrList));

				// If everything goes right, add the new recipient to the message
				// object passed into us.
				EC_MAPI(lpMessage->ModifyRecipients(MODRECIP_ADD, lpAdrList));

				EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			}

			if (lpAdrList) FreePadrlist(lpAdrList);
			if (lpAddrBook) lpAddrBook->Release();
			return hRes;
		}

		// Same as CreatePropertyStringRestriction, but skips the existence part.
		_Check_return_ HRESULT CreateANRRestriction(ULONG ulPropTag,
			_In_ const std::wstring& szString,
			_In_opt_ LPVOID lpParent,
			_Deref_out_opt_ LPSRestriction* lppRes)
		{
			if (PROP_TYPE(ulPropTag) != PT_UNICODE)
			{
				DebugPrint(DBGGeneric, L"Failed to create restriction - property was not PT_UNICODE\n");
				return MAPI_E_INVALID_PARAMETER;
			}

			auto hRes = S_OK;
			LPSRestriction lpRes = nullptr;
			LPSPropValue lpspvSubject = nullptr;
			LPVOID lpAllocationParent = nullptr;

			*lppRes = nullptr;

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
				sizeof(SPropValue),
				lpAllocationParent,
				reinterpret_cast<LPVOID*>(&lpspvSubject)));

			if (lpRes && lpspvSubject)
			{
				// Zero out allocated memory.
				ZeroMemory(lpRes, sizeof(SRestriction));
				ZeroMemory(lpspvSubject, sizeof(SPropValue));

				// Root Node
				lpRes->rt = RES_PROPERTY;
				lpRes->res.resProperty.relop = RELOP_EQ;
				lpRes->res.resProperty.ulPropTag = ulPropTag;
				lpRes->res.resProperty.lpProp = lpspvSubject;

				// Allocate and fill out properties:
				lpspvSubject->ulPropTag = ulPropTag;
				lpspvSubject->Value.LPSZ = nullptr;

				if (!szString.empty())
				{
					lpspvSubject->Value.lpszW = const_cast<LPWSTR>(szString.c_str());
				}

				DebugPrint(DBGGeneric, L"CreateANRRestriction built restriction:\n");
				DebugPrintRestriction(DBGGeneric, lpRes, nullptr);

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

		_Check_return_ HRESULT GetABContainerTable(_In_ LPADRBOOK lpAdrBook, _Deref_out_opt_ LPMAPITABLE* lpABContainerTable)
		{
			auto hRes = S_OK;
			LPABCONT lpABRootContainer = nullptr;
			LPMAPITABLE lpTable = nullptr;

			*lpABContainerTable = nullptr;
			if (!lpAdrBook) return MAPI_E_INVALID_PARAMETER;

			// Open root address book (container).
			EC_H(mapi::CallOpenEntry(
				nullptr,
				lpAdrBook,
				nullptr,
				nullptr,
				nullptr,
				nullptr,
				NULL,
				nullptr,
				reinterpret_cast<LPUNKNOWN*>(&lpABRootContainer)));

			if (lpABRootContainer)
			{
				// Get a table of all of the Address Books.
				EC_MAPI(lpABRootContainer->GetHierarchyTable(CONVENIENT_DEPTH | fMapiUnicode, &lpTable));
				*lpABContainerTable = lpTable;
				lpABRootContainer->Release();
			}

			return hRes;
		}

		// Manually resolve a name in the address book and add it to the message
		_Check_return_ HRESULT ManualResolve(
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMESSAGE lpMessage,
			_In_ const std::wstring& szName,
			ULONG PropTagToCompare)
		{
			auto hRes = S_OK;
			ULONG ulObjType = 0;
			LPADRBOOK lpAdrBook = nullptr;
			LPSRowSet lpABRow = nullptr;
			LPMAPITABLE lpABContainerTable = nullptr;
			LPADRLIST lpAdrList = nullptr;
			LPABCONT lpABContainer = nullptr;
			LPMAPITABLE pTable = nullptr;
			LPSPropValue lpFoundRow = nullptr;

			enum
			{
				abcPR_ENTRYID,
				abcPR_DISPLAY_NAME,
				abcNUM_COLS
			};

			static const SizedSPropTagArray(abcNUM_COLS, abcCols) =
			{
			abcNUM_COLS,
				{
					PR_ENTRYID,
					PR_DISPLAY_NAME
				},
			};

			enum
			{
				abPR_ENTRYID,
				abPR_DISPLAY_NAME,
				abPR_RECIPIENT_TYPE,
				abPR_ADDRTYPE,
				abPR_DISPLAY_TYPE,
				abPropTagToCompare,
				abNUM_COLS
			};

			if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;
			if (PROP_TYPE(PropTagToCompare) != PT_UNICODE) return MAPI_E_INVALID_PARAMETER;

			DebugPrint(DBGGeneric, L"ManualResolve: Asked to resolve \"%ws\"\n", szName.c_str());

			EC_MAPI(lpMAPISession->OpenAddressBook(
				NULL,
				nullptr,
				NULL,
				&lpAdrBook));

			EC_H(GetABContainerTable(lpAdrBook, &lpABContainerTable));

			if (lpABContainerTable)
			{
				// Restrict the table to the properties that we are interested in.
				EC_MAPI(lpABContainerTable->SetColumns(LPSPropTagArray(&abcCols), TBL_BATCH));

				if (!FAILED(hRes)) for (;;)
				{
					hRes = S_OK;

					FreeProws(lpABRow);
					lpABRow = nullptr;
					EC_MAPI(lpABContainerTable->QueryRows(
						1,
						NULL,
						&lpABRow));
					if (FAILED(hRes) || !lpABRow || lpABRow && !lpABRow->cRows) break;

					// From this point forward, consider any error an error with the current address book container, so just continue and try the next one.
					if (PR_ENTRYID == lpABRow->aRow->lpProps[abcPR_ENTRYID].ulPropTag)
					{
						DebugPrint(DBGGeneric, L"ManualResolve: Searching this container\n");
						DebugPrintBinary(DBGGeneric, lpABRow->aRow->lpProps[abcPR_ENTRYID].Value.bin);

						if (lpABContainer) lpABContainer->Release();
						lpABContainer = nullptr;
						EC_H(mapi::CallOpenEntry(
							nullptr,
							lpAdrBook,
							nullptr,
							nullptr,
							lpABRow->aRow->lpProps[abcPR_ENTRYID].Value.bin.cb,
							reinterpret_cast<ENTRYID*>(lpABRow->aRow->lpProps[abcPR_ENTRYID].Value.bin.lpb),
							nullptr,
							NULL,
							&ulObjType,
							reinterpret_cast<LPUNKNOWN*>(&lpABContainer)));
						if (!lpABContainer) continue;

						DebugPrint(DBGGeneric, L"ManualResolve: Object opened as 0x%X\n", ulObjType);

						if (lpABContainer && ulObjType == MAPI_ABCONT)
						{
							if (pTable) pTable->Release();
							pTable = nullptr;
							WC_MAPI(lpABContainer->GetContentsTable(fMapiUnicode, &pTable));
							if (!pTable)
							{
								DebugPrint(DBGGeneric, L"ManualResolve: Container did not support contents table\n");
								continue;
							}

							MAPIFreeBuffer(lpFoundRow);
							lpFoundRow = nullptr;
							EC_H(SearchContentsTableForName(
								pTable,
								szName,
								PropTagToCompare,
								&lpFoundRow));
							if (!lpFoundRow) continue;

							if (lpAdrList) FreePadrlist(lpAdrList);
							lpAdrList = nullptr;
							// Allocate memory for new Address List structure.
							EC_H(MAPIAllocateBuffer(CbNewADRLIST(1), reinterpret_cast<LPVOID*>(&lpAdrList)));
							if (!lpAdrList) continue;

							ZeroMemory(lpAdrList, CbNewADRLIST(1));
							lpAdrList->cEntries = 1;
							// Allocate memory for SPropValue structure that indicates what
							// recipient properties will be set. To resolve a name that
							// already exists in the Address book, this will always be 1.

							EC_H(MAPIAllocateBuffer(
								static_cast<ULONG>(abNUM_COLS * sizeof(SPropValue)),
								reinterpret_cast<LPVOID*>(&lpAdrList->aEntries->rgPropVals)));
							if (!lpAdrList->aEntries->rgPropVals) continue;

							// TODO: We are setting 5 properties below. If this changes, modify these two lines.
							ZeroMemory(lpAdrList->aEntries->rgPropVals, 5 * sizeof(SPropValue));
							lpAdrList->aEntries->cValues = 5;

							// Fill out addresslist with required property values.
							const auto pProps = lpAdrList->aEntries->rgPropVals;

							auto pProp = &pProps[abPR_ENTRYID]; // Just a pointer, do not free.
							pProp->ulPropTag = PR_ENTRYID;
							EC_H(mapi::CopySBinary(&pProp->Value.bin, &lpFoundRow[abPR_ENTRYID].Value.bin, lpAdrList));

							pProp = &pProps[abPR_RECIPIENT_TYPE];
							pProp->ulPropTag = PR_RECIPIENT_TYPE;
							pProp->Value.l = MAPI_TO;

							pProp = &pProps[abPR_DISPLAY_NAME];
							pProp->ulPropTag = PR_DISPLAY_NAME;

							if (!mapi::CheckStringProp(&lpFoundRow[abPR_DISPLAY_NAME], PT_TSTRING)) continue;

							EC_H(mapi::CopyString(
								&pProp->Value.LPSZ,
								lpFoundRow[abPR_DISPLAY_NAME].Value.LPSZ,
								lpAdrList));

							pProp = &pProps[abPR_ADDRTYPE];
							pProp->ulPropTag = PR_ADDRTYPE;

							if (!mapi::CheckStringProp(&lpFoundRow[abPR_ADDRTYPE], PT_TSTRING)) continue;

							EC_H(mapi::CopyString(
								&pProp->Value.LPSZ,
								lpFoundRow[abPR_ADDRTYPE].Value.LPSZ,
								lpAdrList));

							pProp = &pProps[abPR_DISPLAY_TYPE];
							pProp->ulPropTag = PR_DISPLAY_TYPE;
							pProp->Value.l = lpFoundRow[abPR_DISPLAY_TYPE].Value.l;

							EC_MAPI(lpMessage->ModifyRecipients(
								MODRECIP_ADD,
								lpAdrList));

							if (lpAdrList) FreePadrlist(lpAdrList);
							lpAdrList = nullptr;

							EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

							// since we're done with our work, let's get out of here.
							break;
						}
					}
				}
				lpABContainerTable->Release();
			}
			FreeProws(lpABRow);
			MAPIFreeBuffer(lpFoundRow);
			if (lpAdrList) FreePadrlist(lpAdrList);
			if (pTable) pTable->Release();
			if (lpABContainer) lpABContainer->Release();
			if (lpAdrBook) lpAdrBook->Release();
			return hRes;
		}

		_Check_return_ HRESULT SearchContentsTableForName(
			_In_ LPMAPITABLE pTable,
			_In_ const std::wstring& szName,
			ULONG PropTagToCompare,
			_Deref_out_opt_ LPSPropValue *lppPropsFound)
		{
			auto hRes = S_OK;

			LPSRowSet pRows = nullptr;

			enum
			{
				abPR_ENTRYID,
				abPR_DISPLAY_NAME,
				abPR_RECIPIENT_TYPE,
				abPR_ADDRTYPE,
				abPR_DISPLAY_TYPE,
				abPropTagToCompare,
				abNUM_COLS
			};

			const SizedSPropTagArray(abNUM_COLS, abCols) =
			{
			abNUM_COLS,
				{
					PR_ENTRYID,
					PR_DISPLAY_NAME_W,
					PR_RECIPIENT_TYPE,
					PR_ADDRTYPE_W,
					PR_DISPLAY_TYPE,
					PropTagToCompare
				}
			};

			*lppPropsFound = nullptr;
			if (!pTable || szName.empty()) return MAPI_E_INVALID_PARAMETER;
			if (PROP_TYPE(PropTagToCompare) != PT_UNICODE) return MAPI_E_INVALID_PARAMETER;

			DebugPrint(DBGGeneric, L"SearchContentsTableForName: Looking for \"%ws\"\n", szName.c_str());

			// Set a restriction so we only find close matches
			LPSRestriction lpSRes = nullptr;

			EC_H(CreateANRRestriction(
				PR_ANR_W,
				szName,
				nullptr,
				&lpSRes));

			EC_MAPI(pTable->SetColumns(LPSPropTagArray(&abCols), TBL_BATCH));

			// Jump to the top of the table...
			EC_MAPI(pTable->SeekRow(
				BOOKMARK_BEGINNING,
				0,
				nullptr));

			// ..and jump to the first matching entry in the table
			EC_MAPI(pTable->Restrict(
				lpSRes,
				NULL
			));

			// Now we iterate through each of the matching entries
			if (!FAILED(hRes)) for (;;)
			{
				hRes = S_OK;
				if (pRows) FreeProws(pRows);
				pRows = nullptr;
				EC_MAPI(pTable->QueryRows(
					1,
					NULL,
					&pRows));
				if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

				// An error at this point is an error with the current entry, so we can continue this for statement
				// Unless it's an allocation error. Those are bad.
				if (PropTagToCompare == pRows->aRow->lpProps[abPropTagToCompare].ulPropTag &&
					mapi::CheckStringProp(&pRows->aRow->lpProps[abPropTagToCompare], PT_UNICODE))
				{
					DebugPrint(DBGGeneric, L"SearchContentsTableForName: found \"%ws\"\n", pRows->aRow->lpProps[abPropTagToCompare].Value.lpszW);
					if (szName == pRows->aRow->lpProps[abPropTagToCompare].Value.lpszW)
					{
						DebugPrint(DBGGeneric, L"SearchContentsTableForName: This is an exact match!\n");
						// We found a match! Return it!
						EC_MAPI(ScDupPropset(
							abNUM_COLS,
							pRows->aRow->lpProps,
							MAPIAllocateBuffer,
							lppPropsFound));
						break;
					}
				}
			}

			if (!*lppPropsFound)
			{
				DebugPrint(DBGGeneric, L"SearchContentsTableForName: No exact match found.\n");
			}
			MAPIFreeBuffer(lpSRes);
			if (pRows) FreeProws(pRows);
			return hRes;
		}

		_Check_return_ HRESULT SelectUser(_In_ LPADRBOOK lpAdrBook, HWND hwnd, _Out_opt_ ULONG* lpulObjType, _Deref_out_opt_ LPMAILUSER* lppMailUser)
		{
			if (!lpAdrBook || !hwnd || !lppMailUser) return MAPI_E_INVALID_PARAMETER;

			auto hRes = S_OK;

			ADRPARM AdrParm = { 0 };
			LPADRLIST lpAdrList = nullptr;
			LPSPropValue lpEntryID = nullptr;
			LPMAILUSER lpMailUser = nullptr;

			*lppMailUser = nullptr;
			if (lpulObjType)
			{
				*lpulObjType = NULL;
			}

			auto szTitle = strings::wstringTostring(strings::loadstring(IDS_SELECTMAILBOX));

			AdrParm.ulFlags = DIALOG_MODAL | ADDRESS_ONE | AB_SELECTONLY | AB_RESOLVE;
			AdrParm.lpszCaption = LPTSTR(szTitle.c_str());

			EC_H_CANCEL(lpAdrBook->Address(
				reinterpret_cast<ULONG_PTR*>(&hwnd),
				&AdrParm,
				&lpAdrList));

			if (lpAdrList)
			{
				lpEntryID = PpropFindProp(
					lpAdrList[0].aEntries->rgPropVals,
					lpAdrList[0].aEntries->cValues,
					PR_ENTRYID);

				if (lpEntryID)
				{
					ULONG ulObjType = NULL;

					EC_H(mapi::CallOpenEntry(
						nullptr,
						lpAdrBook,
						nullptr,
						nullptr,
						lpEntryID->Value.bin.cb,
						reinterpret_cast<LPENTRYID>(lpEntryID->Value.bin.lpb),
						nullptr,
						MAPI_BEST_ACCESS,
						&ulObjType,
						reinterpret_cast<LPUNKNOWN*>(&lpMailUser)));
					if (SUCCEEDED(hRes) && lpMailUser)
					{
						*lppMailUser = lpMailUser;
					}
					else
					{
						if (lpMailUser)
						{
							lpMailUser->Release();
						}

						if (lpulObjType)
						{
							*lpulObjType = ulObjType;
						}
					}
				}
			}

			if (lpAdrList) FreePadrlist(lpAdrList);
			return hRes;
		}
	}
}