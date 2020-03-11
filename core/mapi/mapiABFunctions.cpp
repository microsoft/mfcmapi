// Collection of useful MAPI Address Book functions

#include <core/stdafx.h>
#include <core/mapi/mapiABFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiOutput.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/error.h>
#include <core/mapi/extraPropTags.h>

namespace mapi::ab
{
	_Check_return_ LPADRLIST AllocAdrList(ULONG ulNumProps)
	{
		if (ulNumProps > ULONG_MAX / sizeof(SPropValue)) return nullptr;

		// Allocate memory for new SRowSet structure.
		auto lpLocalAdrList = mapi::allocate<LPADRLIST>(CbNewSRowSet(1));
		if (lpLocalAdrList)
		{
			// Allocate memory for SPropValue structure that indicates what recipient properties will be set.
			lpLocalAdrList->aEntries[0].rgPropVals = mapi::allocate<LPSPropValue>(ulNumProps * sizeof(SPropValue));
			if (!lpLocalAdrList->aEntries[0].rgPropVals)
			{
				FreePadrlist(lpLocalAdrList);
				lpLocalAdrList = nullptr;
			}
		}

		return lpLocalAdrList;
	}

	_Check_return_ HRESULT AddOneOffAddress(
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMESSAGE lpMessage,
		_In_ const std::wstring& szDisplayName,
		_In_ const std::wstring& szAddrType,
		_In_ const std::wstring& szEmailAddress,
		ULONG ulRecipientType)
	{
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

		auto hRes = EC_MAPI(lpMAPISession->OpenAddressBook(NULL, nullptr, NULL, &lpAddrBook));

		if (SUCCEEDED(hRes))
		{
			lpAdrList = AllocAdrList(NUM_RECIP_PROPS);
		}

		// Setup the One Time recipient by indicating how many recipients
		// and how many properties will be set on each recipient.
		if (lpAdrList)
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

			auto bin = mapi::setBin(lpAdrList->aEntries[0].rgPropVals[EID]);
			// Create the One-off address and get an EID for it.
			hRes = EC_MAPI(lpAddrBook->CreateOneOff(
				reinterpret_cast<LPTSTR>(lpAdrList->aEntries[0].rgPropVals[NAME].Value.lpszW),
				reinterpret_cast<LPTSTR>(lpAdrList->aEntries[0].rgPropVals[ADDR].Value.lpszW),
				reinterpret_cast<LPTSTR>(lpAdrList->aEntries[0].rgPropVals[EMAIL].Value.lpszW),
				MAPI_UNICODE,
				&bin.cb,
				reinterpret_cast<LPENTRYID*>(&bin.lpb)));

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpAddrBook->ResolveName(0L, MAPI_UNICODE, nullptr, lpAdrList));
			}

			// If everything goes right, add the new recipient to the message
			// object passed into us.
			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpMessage->ModifyRecipients(MODRECIP_ADD, lpAdrList));
			}

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			}
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
		LPADRLIST lpAdrList = nullptr; // ModifyRecips takes LPADRLIST
		LPADRBOOK lpAddrBook = nullptr;

		enum
		{
			NAME,
			RECIP,
			NUM_RECIP_PROPS
		};

		if (!lpMessage || !lpMAPISession) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(lpMAPISession->OpenAddressBook(NULL, nullptr, NULL, &lpAddrBook));

		if (SUCCEEDED(hRes))
		{
			lpAdrList = AllocAdrList(NUM_RECIP_PROPS);
		}

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

			hRes = EC_MAPI(lpAddrBook->ResolveName(0L, MAPI_UNICODE, nullptr, lpAdrList));

			if (SUCCEEDED(hRes))
			{ // If everything goes right, add the new recipient to the message
				// object passed into us.
				hRes = EC_MAPI(lpMessage->ModifyRecipients(MODRECIP_ADD, lpAdrList));
			}

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			}
		}

		if (lpAdrList) FreePadrlist(lpAdrList);
		if (lpAddrBook) lpAddrBook->Release();
		return hRes;
	}

	// Same as CreatePropertyStringRestriction, but skips the existence part.
	_Check_return_ LPSRestriction
	CreateANRRestriction(ULONG ulPropTag, _In_ const std::wstring& szString, _In_opt_ LPVOID lpParent)
	{
		if (PROP_TYPE(ulPropTag) != PT_UNICODE)
		{
			output::DebugPrint(
				output::dbgLevel::Generic, L"Failed to create restriction - property was not PT_UNICODE\n");
			return nullptr;
		}

		// Allocate and create our SRestriction
		// Allocate base memory:
		const auto lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), lpParent);
		const auto lpAllocationParent = lpParent ? lpParent : lpRes;

		if (lpRes)
		{
			// Root Node
			lpRes->rt = RES_PROPERTY;
			lpRes->res.resProperty.relop = RELOP_EQ;
			lpRes->res.resProperty.ulPropTag = ulPropTag;
			const auto lpspvSubject = mapi::allocate<LPSPropValue>(sizeof(SPropValue), lpAllocationParent);
			lpRes->res.resProperty.lpProp = lpspvSubject;

			if (lpspvSubject)
			{ // Allocate and fill out properties:
				lpspvSubject->ulPropTag = ulPropTag;
				lpspvSubject->Value.LPSZ = nullptr;

				if (!szString.empty())
				{
					lpspvSubject->Value.lpszW = CopyStringW(szString.c_str(), lpAllocationParent);
				}
			}

			output::DebugPrint(output::dbgLevel::Generic, L"CreateANRRestriction built restriction:\n");
			output::outputRestriction(output::dbgLevel::Generic, nullptr, lpRes, nullptr);
		}

		return lpRes;
	}

	_Check_return_ LPMAPITABLE GetABContainerTable(_In_ LPADRBOOK lpAdrBook)
	{
		if (!lpAdrBook) return nullptr;

		LPMAPITABLE lpTable = nullptr;

		// Open root address book (container).
		auto lpABRootContainer =
			mapi::CallOpenEntry<LPABCONT>(nullptr, lpAdrBook, nullptr, nullptr, nullptr, nullptr, NULL, nullptr);

		if (lpABRootContainer)
		{
			// Get a table of all of the Address Books.
			EC_MAPI_S(lpABRootContainer->GetHierarchyTable(CONVENIENT_DEPTH | fMapiUnicode, &lpTable));
			lpABRootContainer->Release();
		}

		return lpTable;
	}

	// Manually resolve a name in the address book and add it to the message
	_Check_return_ HRESULT ManualResolve(
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMESSAGE lpMessage,
		_In_ const std::wstring& szName,
		ULONG PropTagToCompare)
	{
		ULONG ulObjType = 0;
		LPADRBOOK lpAdrBook = nullptr;
		LPSRowSet lpABRow = nullptr;
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

		static const SizedSPropTagArray(abcNUM_COLS, abcCols) = {
			abcNUM_COLS,
			{
				PR_ENTRYID, //
				PR_DISPLAY_NAME_W //
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

		output::DebugPrint(output::dbgLevel::Generic, L"ManualResolve: Asked to resolve \"%ws\"\n", szName.c_str());

		auto hRes = EC_MAPI(lpMAPISession->OpenAddressBook(NULL, nullptr, NULL, &lpAdrBook));

		LPMAPITABLE lpABContainerTable = nullptr;
		if (SUCCEEDED(hRes))
		{
			lpABContainerTable = GetABContainerTable(lpAdrBook);
		}

		if (lpABContainerTable)
		{
			// Restrict the table to the properties that we are interested in.
			hRes = EC_MAPI(lpABContainerTable->SetColumns(LPSPropTagArray(&abcCols), TBL_BATCH));

			if (SUCCEEDED(hRes))
			{
				for (;;)
				{
					FreeProws(lpABRow);
					lpABRow = nullptr;
					hRes = EC_MAPI(lpABContainerTable->QueryRows(1, NULL, &lpABRow));
					if (FAILED(hRes) || !lpABRow || lpABRow && !lpABRow->cRows) break;

					// From this point forward, consider any error an error with the current address book container, so just continue and try the next one.
					if (PR_ENTRYID == lpABRow->aRow->lpProps[abcPR_ENTRYID].ulPropTag)
					{
						auto bin = mapi::getBin(lpABRow->aRow->lpProps[abcPR_ENTRYID]);
						output::DebugPrint(output::dbgLevel::Generic, L"ManualResolve: Searching this container\n");
						output::outputBinary(output::dbgLevel::Generic, nullptr, bin);

						if (lpABContainer) lpABContainer->Release();
						lpABContainer = mapi::CallOpenEntry<LPABCONT>(
							nullptr,
							lpAdrBook,
							nullptr,
							nullptr,
							bin.cb,
							reinterpret_cast<ENTRYID*>(bin.lpb),
							nullptr,
							NULL,
							&ulObjType);
						if (!lpABContainer) continue;

						output::DebugPrint(
							output::dbgLevel::Generic, L"ManualResolve: Object opened as 0x%X\n", ulObjType);

						if (lpABContainer && ulObjType == MAPI_ABCONT)
						{
							if (pTable) pTable->Release();
							pTable = nullptr;
							WC_MAPI_S(lpABContainer->GetContentsTable(fMapiUnicode, &pTable));
							if (!pTable)
							{
								output::DebugPrint(
									output::dbgLevel::Generic,
									L"ManualResolve: Container did not support contents table\n");
								continue;
							}

							MAPIFreeBuffer(lpFoundRow);
							lpFoundRow = nullptr;
							hRes = EC_H(SearchContentsTableForName(pTable, szName, PropTagToCompare, &lpFoundRow));
							if (!lpFoundRow) continue;

							if (lpAdrList) FreePadrlist(lpAdrList);
							// Allocate memory for new Address List structure.
							lpAdrList = mapi::allocate<LPADRLIST>(CbNewADRLIST(1));
							if (!lpAdrList) continue;

							lpAdrList->cEntries = 1;
							// Allocate memory for SPropValue structure that indicates what
							// recipient properties will be set. To resolve a name that
							// already exists in the Address book, this will always be 1.

							lpAdrList->aEntries->rgPropVals =
								mapi::allocate<LPSPropValue>(ULONG{abNUM_COLS * sizeof(SPropValue)});
							if (!lpAdrList->aEntries->rgPropVals) continue;

							// We are setting 5 properties below. If this changes, modify these two lines.
							lpAdrList->aEntries->cValues = 5;

							// Fill out addresslist with required property values.
							const auto pProps = lpAdrList->aEntries->rgPropVals;

							auto pProp = &pProps[abPR_ENTRYID]; // Just a pointer, do not free.
							pProp->ulPropTag = PR_ENTRYID;
							mapi::setBin(pProp) = CopySBinary(mapi::getBin(lpFoundRow[abPR_ENTRYID]), lpAdrList);

							pProp = &pProps[abPR_RECIPIENT_TYPE];
							pProp->ulPropTag = PR_RECIPIENT_TYPE;
							pProp->Value.l = MAPI_TO;

							pProp = &pProps[abPR_DISPLAY_NAME];
							pProp->ulPropTag = PR_DISPLAY_NAME_W;
							if (!strings::CheckStringProp(&lpFoundRow[abPR_DISPLAY_NAME], PT_UNICODE)) continue;
							pProp->Value.lpszW = CopyStringW(lpFoundRow[abPR_DISPLAY_NAME].Value.lpszW, lpAdrList);

							pProp = &pProps[abPR_ADDRTYPE];
							pProp->ulPropTag = PR_ADDRTYPE_W;
							if (!strings::CheckStringProp(&lpFoundRow[abPR_ADDRTYPE], PT_UNICODE)) continue;
							pProp->Value.lpszW = CopyStringW(lpFoundRow[abPR_ADDRTYPE].Value.lpszW, lpAdrList);

							pProp = &pProps[abPR_DISPLAY_TYPE];
							pProp->ulPropTag = PR_DISPLAY_TYPE;
							pProp->Value.l = lpFoundRow[abPR_DISPLAY_TYPE].Value.l;

							if (SUCCEEDED(hRes))
							{
								hRes = EC_MAPI(lpMessage->ModifyRecipients(MODRECIP_ADD, lpAdrList));
							}

							if (lpAdrList) FreePadrlist(lpAdrList);
							lpAdrList = nullptr;

							if (SUCCEEDED(hRes))
							{
								hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
							}

							// since we're done with our work, let's get out of here.
							break;
						}
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
		_Deref_out_opt_ LPSPropValue* lppPropsFound)
	{
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

		const SizedSPropTagArray(abNUM_COLS, abCols) = {
			abNUM_COLS,
			{
				PR_ENTRYID, //
				PR_DISPLAY_NAME_W, //
				PR_RECIPIENT_TYPE, //
				PR_ADDRTYPE_W, //
				PR_DISPLAY_TYPE, //
				PropTagToCompare //
			},
		};

		*lppPropsFound = nullptr;
		if (!pTable || szName.empty()) return MAPI_E_INVALID_PARAMETER;
		if (PROP_TYPE(PropTagToCompare) != PT_UNICODE) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(
			output::dbgLevel::Generic, L"SearchContentsTableForName: Looking for \"%ws\"\n", szName.c_str());

		auto hRes = EC_MAPI(pTable->SetColumns(LPSPropTagArray(&abCols), TBL_BATCH));
		if (SUCCEEDED(hRes))
		{
			// Jump to the top of the table...
			hRes = EC_MAPI(pTable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));
		}

		// Set a restriction so we only find close matches
		if (SUCCEEDED(hRes))
		{
			// ..and jump to the first matching entry in the table
			auto lpSRes = CreateANRRestriction(PR_ANR_W, szName, nullptr);
			hRes = EC_MAPI(pTable->Restrict(lpSRes, NULL));
			MAPIFreeBuffer(lpSRes);
		}

		// Now we iterate through each of the matching entries
		if (SUCCEEDED(hRes))
		{
			for (;;)
			{
				if (pRows) FreeProws(pRows);
				pRows = nullptr;
				hRes = EC_MAPI(pTable->QueryRows(1, NULL, &pRows));
				if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

				// An error at this point is an error with the current entry, so we can continue this for statement
				// Unless it's an allocation error. Those are bad.
				if (PropTagToCompare == pRows->aRow->lpProps[abPropTagToCompare].ulPropTag &&
					strings::CheckStringProp(&pRows->aRow->lpProps[abPropTagToCompare], PT_UNICODE))
				{
					output::DebugPrint(
						output::dbgLevel::Generic,
						L"SearchContentsTableForName: found \"%ws\"\n",
						pRows->aRow->lpProps[abPropTagToCompare].Value.lpszW);
					if (szName == pRows->aRow->lpProps[abPropTagToCompare].Value.lpszW)
					{
						output::DebugPrint(
							output::dbgLevel::Generic, L"SearchContentsTableForName: This is an exact match!\n");
						// We found a match! Return it!
						hRes =
							EC_MAPI(ScDupPropset(abNUM_COLS, pRows->aRow->lpProps, MAPIAllocateBuffer, lppPropsFound));
						break;
					}
				}
			}
		}

		if (!*lppPropsFound)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"SearchContentsTableForName: No exact match found.\n");
		}

		if (pRows) FreeProws(pRows);
		return hRes;
	}

	_Check_return_ LPMAILUSER SelectUser(_In_ LPADRBOOK lpAdrBook, HWND hwnd, _Out_opt_ ULONG* lpulObjType)
	{
		if (lpulObjType)
		{
			*lpulObjType = NULL;
		}

		if (!lpAdrBook || !hwnd) return nullptr;

		auto szTitle = strings::wstringTotstring(strings::loadstring(IDS_SELECTMAILBOX));

		ADRPARM AdrParm = {};
		AdrParm.ulFlags = DIALOG_MODAL | ADDRESS_ONE | AB_SELECTONLY | AB_RESOLVE;
#ifdef _UNICODE
		AdrParm.ulFlags |= AB_UNICODEUI;
#endif
		AdrParm.lpszCaption = reinterpret_cast<LPTSTR>(const_cast<LPSTR>(szTitle.c_str()));

		LPMAILUSER lpMailUser = nullptr;
		LPADRLIST lpAdrList = nullptr;
		const auto hRes = EC_H_CANCEL(lpAdrBook->Address(reinterpret_cast<ULONG_PTR*>(&hwnd), &AdrParm, &lpAdrList));
		if (lpAdrList)
		{
			const auto lpEntryID =
				PpropFindProp(lpAdrList[0].aEntries->rgPropVals, lpAdrList[0].aEntries->cValues, PR_ENTRYID);

			if (lpEntryID)
			{
				ULONG ulObjType = NULL;

				auto bin = mapi::getBin(lpEntryID);
				lpMailUser = mapi::CallOpenEntry<LPMAILUSER>(
					nullptr,
					lpAdrBook,
					nullptr,
					nullptr,
					bin.cb,
					reinterpret_cast<LPENTRYID>(bin.lpb),
					nullptr,
					MAPI_BEST_ACCESS,
					&ulObjType);
				if (FAILED(hRes))
				{
					if (lpMailUser)
					{
						lpMailUser->Release();
						lpMailUser = nullptr;
					}
				}

				if (lpulObjType)
				{
					*lpulObjType = ulObjType;
				}
			}

			FreePadrlist(lpAdrList);
		}

		return lpMailUser;
	}
} // namespace mapi::ab