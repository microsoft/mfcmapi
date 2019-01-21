#include <StdAfx.h>
#include <IO/MFCOutput.h>
#include <core/utility/output.h>
#include <MAPI/MAPIFunctions.h>
#include <core/utility/strings.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/SmartView/SmartView.h>
#include <core/mapi/columnTags.h>
#include <Property/ParseProperty.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/registry.h>
#include <cassert>
#include <core/interpret/flags.h>
#include <core/interpret/proptags.h>
#include <core/interpret/proptype.h>

#ifdef CHECKFORMATPARAMS
#undef Outputf
#undef OutputToFilef
#undef DebugPrint
#undef DebugPrintEx
#endif

namespace output
{
	void _OutputNamedPropID(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPINAMEID lpName)
	{
		CHKPARAM;
		EARLYABORT;

		if (!lpName) return;

		if (lpName->ulKind == MNID_ID)
		{
			Outputf(
				ulDbgLvl,
				fFile,
				true,
				L"\t\t: nmid ID: 0x%X\n", // STRING_OK
				lpName->Kind.lID);
		}
		else
		{
			Outputf(
				ulDbgLvl,
				fFile,
				true,
				L"\t\t: nmid Name: %ws\n", // STRING_OK
				lpName->Kind.lpwstrName);
		}

		Output(ulDbgLvl, fFile, false, guid::GUIDToStringAndName(lpName->lpguid));
		Output(ulDbgLvl, fFile, false, L"\n");
	}

	void _OutputFormInfo(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMINFO lpMAPIFormInfo)
	{
		CHKPARAM;
		EARLYABORT;
		if (!lpMAPIFormInfo) return;

		LPSPropValue lpPropVals = nullptr;
		ULONG ulPropVals = NULL;
		LPMAPIVERBARRAY lpMAPIVerbArray = nullptr;
		LPMAPIFORMPROPARRAY lpMAPIFormPropArray = nullptr;

		Outputf(ulDbgLvl, fFile, true, L"Dumping verb and property set for form: %p\n", lpMAPIFormInfo);

		EC_H_S(mapi::GetPropsNULL(lpMAPIFormInfo, fMapiUnicode, &ulPropVals, &lpPropVals));
		if (lpPropVals)
		{
			_OutputProperties(ulDbgLvl, fFile, ulPropVals, lpPropVals, lpMAPIFormInfo, false);
			MAPIFreeBuffer(lpPropVals);
		}

		EC_MAPI_S(lpMAPIFormInfo->CalcVerbSet(NULL, &lpMAPIVerbArray)); // API doesn't support Unicode
		if (lpMAPIVerbArray)
		{
			Outputf(ulDbgLvl, fFile, true, L"\t0x%X verbs:\n", lpMAPIVerbArray->cMAPIVerb);
			for (ULONG i = 0; i < lpMAPIVerbArray->cMAPIVerb; i++)
			{
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"\t\tVerb 0x%X\n", // STRING_OK
					i);
				if (lpMAPIVerbArray->aMAPIVerb[i].ulFlags == MAPI_UNICODE)
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"\t\tDoVerb value: 0x%X\n\t\tUnicode Name: %ws\n\t\tFlags: 0x%X\n\t\tAttributes: 0x%X\n", // STRING_OK
						lpMAPIVerbArray->aMAPIVerb[i].lVerb,
						LPWSTR(lpMAPIVerbArray->aMAPIVerb[i].szVerbname),
						lpMAPIVerbArray->aMAPIVerb[i].fuFlags,
						lpMAPIVerbArray->aMAPIVerb[i].grfAttribs);
				}
				else
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"\t\tDoVerb value: 0x%X\n\t\tANSI Name: %hs\n\t\tFlags: 0x%X\n\t\tAttributes: 0x%X\n", // STRING_OK
						lpMAPIVerbArray->aMAPIVerb[i].lVerb,
						LPSTR(lpMAPIVerbArray->aMAPIVerb[i].szVerbname),
						lpMAPIVerbArray->aMAPIVerb[i].fuFlags,
						lpMAPIVerbArray->aMAPIVerb[i].grfAttribs);
				}
			}

			MAPIFreeBuffer(lpMAPIVerbArray);
		}

		EC_MAPI_S(lpMAPIFormInfo->CalcFormPropSet(NULL, &lpMAPIFormPropArray)); // API doesn't support Unicode
		if (lpMAPIFormPropArray)
		{
			_OutputFormPropArray(ulDbgLvl, fFile, lpMAPIFormPropArray);
			MAPIFreeBuffer(lpMAPIFormPropArray);
		}
	}

	void _OutputFormPropArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMPROPARRAY lpMAPIFormPropArray)
	{
		Outputf(ulDbgLvl, fFile, true, L"\t0x%X Properties:\n", lpMAPIFormPropArray->cProps);
		for (ULONG i = 0; i < lpMAPIFormPropArray->cProps; i++)
		{
			Outputf(ulDbgLvl, fFile, true, L"\t\tProperty 0x%X\n", i);

			if (lpMAPIFormPropArray->aFormProp[i].ulFlags == MAPI_UNICODE)
			{
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"\t\tProperty Name: %ws\n\t\tProperty Type: %ws\n\t\tSpecial Type: 0x%X\n\t\tNum Vals: 0x%X\n", // STRING_OK
					LPWSTR(lpMAPIFormPropArray->aFormProp[i].pszDisplayName),
					proptype::TypeToString(lpMAPIFormPropArray->aFormProp[i].nPropType).c_str(),
					lpMAPIFormPropArray->aFormProp[i].nSpecialType,
					lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable);
			}
			else
			{
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"\t\tProperty Name: %hs\n\t\tProperty Type: %ws\n\t\tSpecial Type: 0x%X\n\t\tNum Vals: 0x%X\n", // STRING_OK
					LPSTR(lpMAPIFormPropArray->aFormProp[i].pszDisplayName),
					proptype::TypeToString(lpMAPIFormPropArray->aFormProp[i].nPropType).c_str(),
					lpMAPIFormPropArray->aFormProp[i].nSpecialType,
					lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable);
			}

			_OutputNamedPropID(ulDbgLvl, fFile, &lpMAPIFormPropArray->aFormProp[i].u.s1.nmidIdx);
			for (ULONG j = 0; j < lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable; j++)
			{
				if (lpMAPIFormPropArray->aFormProp[i].ulFlags == MAPI_UNICODE)
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"\t\t\tEnum 0x%X\nEnumVal Name: %ws\t\t\t\nEnumVal enumeration: 0x%X\n", // STRING_OK
						j,
						LPWSTR(lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].pszDisplayName),
						lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].nVal);
				}
				else
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"\t\t\tEnum 0x%X\nEnumVal Name: %hs\t\t\t\nEnumVal enumeration: 0x%X\n", // STRING_OK
						j,
						LPSTR(lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].pszDisplayName),
						lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].nVal);
				}
			}

			_OutputNamedPropID(ulDbgLvl, fFile, &lpMAPIFormPropArray->aFormProp[i].nmid);
		}
	}

	void _OutputPropTagArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropTagArray lpTagsToDump)
	{
		CHKPARAM;
		EARLYABORT;
		if (!lpTagsToDump) return;

		Outputf(
			ulDbgLvl,
			fFile,
			true,
			L"\tProp tag list, %u props\n", // STRING_OK
			lpTagsToDump->cValues);
		for (ULONG uCurProp = 0; uCurProp < lpTagsToDump->cValues; uCurProp++)
		{
			Outputf(
				ulDbgLvl,
				fFile,
				true,
				L"\t\tProp: %u = %ws\n", // STRING_OK
				uCurProp,
				proptags::TagToString(lpTagsToDump->aulPropTag[uCurProp], nullptr, false, true).c_str());
		}

		Output(ulDbgLvl, fFile, true, L"\tEnd Prop Tag List\n");
	}

	void _OutputTable(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPITABLE lpMAPITable)
	{
		CHKPARAM;
		EARLYABORT;
		if (!lpMAPITable) return;

		LPSRowSet lpRows = nullptr;

		EC_MAPI_S(lpMAPITable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));

		Output(ulDbgLvl, fFile, false, g_szXMLHeader);
		Output(ulDbgLvl, fFile, false, L"<table>\n");

		for (;;)
		{
			FreeProws(lpRows);
			lpRows = nullptr;
			const auto hRes = EC_MAPI(lpMAPITable->QueryRows(20, NULL, &lpRows));
			if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

			for (ULONG iCurRow = 0; iCurRow < lpRows->cRows; iCurRow++)
			{
				Outputf(ulDbgLvl, fFile, false, L"<row index = \"0x%08X\">\n", iCurRow);
				_OutputSRow(ulDbgLvl, fFile, &lpRows->aRow[iCurRow], nullptr);
				Output(ulDbgLvl, fFile, false, L"</row>\n");
			}
		}

		Output(ulDbgLvl, fFile, false, L"</table>\n");

		FreeProws(lpRows);
		lpRows = nullptr;
	}

	void _OutputNotifications(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		ULONG cNotify,
		_In_count_(cNotify) LPNOTIFICATION lpNotifications,
		_In_opt_ LPMAPIPROP lpObj)
	{
		CHKPARAM;
		EARLYABORT;
		if (!lpNotifications) return;

		Outputf(ulDbgLvl, fFile, true, L"Dumping %u notifications.\n", cNotify);

		std::wstring szFlags;
		std::wstring szPropNum;

		for (ULONG i = 0; i < cNotify; i++)
		{
			Outputf(
				ulDbgLvl, fFile, true, L"lpNotifications[%u].ulEventType = 0x%08X", i, lpNotifications[i].ulEventType);
			szFlags = flags::InterpretFlags(flagNotifEventType, lpNotifications[i].ulEventType);
			if (!szFlags.empty())
			{
				Outputf(ulDbgLvl, fFile, false, L" = %ws", szFlags.c_str());
			}

			Output(ulDbgLvl, fFile, false, L"\n");

			SBinary sbin = {};
			switch (lpNotifications[i].ulEventType)
			{
			case fnevCriticalError:
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.err.ulFlags = 0x%08X\n",
					i,
					lpNotifications[i].info.err.ulFlags);
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.err.scode = 0x%08X\n",
					i,
					lpNotifications[i].info.err.scode);
				sbin.cb = lpNotifications[i].info.err.cbEntryID;
				sbin.lpb = reinterpret_cast<LPBYTE>(lpNotifications[i].info.err.lpEntryID);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.err.lpEntryID = ", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);
				if (lpNotifications[i].info.err.lpMAPIError)
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"lpNotifications[%u].info.err.lpMAPIError = %s\n",
						i,
						interpretprop::MAPIErrToString(
							lpNotifications[i].info.err.ulFlags, *lpNotifications[i].info.err.lpMAPIError)
							.c_str());
				}

				break;
			case fnevExtended:
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.ext.ulEvent = 0x%08X\n",
					i,
					lpNotifications[i].info.ext.ulEvent);
				sbin.cb = lpNotifications[i].info.ext.cb;
				sbin.lpb = lpNotifications[i].info.ext.pbEventParameters;
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.ext.pbEventParameters = \n", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);
				break;
			case fnevNewMail:
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.newmail.ulFlags = 0x%08X\n",
					i,
					lpNotifications[i].info.newmail.ulFlags);
				sbin.cb = lpNotifications[i].info.newmail.cbEntryID;
				sbin.lpb = reinterpret_cast<LPBYTE>(lpNotifications[i].info.newmail.lpEntryID);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.newmail.lpEntryID = \n", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);
				sbin.cb = lpNotifications[i].info.newmail.cbParentID;
				sbin.lpb = reinterpret_cast<LPBYTE>(lpNotifications[i].info.newmail.lpParentID);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.newmail.lpParentID = \n", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);

				if (lpNotifications[i].info.newmail.ulFlags & MAPI_UNICODE)
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"lpNotifications[%u].info.newmail.lpszMessageClass = \"%ws\"\n",
						i,
						reinterpret_cast<LPWSTR>(lpNotifications[i].info.newmail.lpszMessageClass));
				}
				else
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"lpNotifications[%u].info.newmail.lpszMessageClass = \"%hs\"\n",
						i,
						LPSTR(lpNotifications[i].info.newmail.lpszMessageClass));
				}

				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.newmail.ulMessageFlags = 0x%08X",
					i,
					lpNotifications[i].info.newmail.ulMessageFlags);
				szPropNum = smartview::InterpretNumberAsStringProp(
					lpNotifications[i].info.newmail.ulMessageFlags, PR_MESSAGE_FLAGS);
				if (!szPropNum.empty())
				{
					Outputf(ulDbgLvl, fFile, false, L" = %ws", szPropNum.c_str());
				}

				Outputf(ulDbgLvl, fFile, false, L"\n");
				break;
			case fnevTableModified:
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.tab.ulTableEvent = 0x%08X",
					i,
					lpNotifications[i].info.tab.ulTableEvent);
				szFlags = flags::InterpretFlags(flagTableEventType, lpNotifications[i].info.tab.ulTableEvent);
				if (!szFlags.empty())
				{
					Outputf(ulDbgLvl, fFile, false, L" = %ws", szFlags.c_str());
				}

				Outputf(ulDbgLvl, fFile, false, L"\n");

				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.tab.hResult = 0x%08X\n",
					i,
					lpNotifications[i].info.tab.hResult);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.tab.propIndex = \n", i);
				_OutputProperty(ulDbgLvl, fFile, &lpNotifications[i].info.tab.propIndex, lpObj, false);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.tab.propPrior = \n", i);
				_OutputProperty(ulDbgLvl, fFile, &lpNotifications[i].info.tab.propPrior, nullptr, false);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.tab.row = \n", i);
				_OutputSRow(ulDbgLvl, fFile, &lpNotifications[i].info.tab.row, lpObj);
				break;
			case fnevObjectCopied:
			case fnevObjectCreated:
			case fnevObjectDeleted:
			case fnevObjectModified:
			case fnevObjectMoved:
			case fnevSearchComplete:
				sbin.cb = lpNotifications[i].info.obj.cbOldID;
				sbin.lpb = reinterpret_cast<LPBYTE>(lpNotifications[i].info.obj.lpOldID);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.obj.lpOldID = \n", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);
				sbin.cb = lpNotifications[i].info.obj.cbOldParentID;
				sbin.lpb = reinterpret_cast<LPBYTE>(lpNotifications[i].info.obj.lpOldParentID);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.obj.lpOldParentID = \n", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);
				sbin.cb = lpNotifications[i].info.obj.cbEntryID;
				sbin.lpb = reinterpret_cast<LPBYTE>(lpNotifications[i].info.obj.lpEntryID);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.obj.lpEntryID = \n", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);
				sbin.cb = lpNotifications[i].info.obj.cbParentID;
				sbin.lpb = reinterpret_cast<LPBYTE>(lpNotifications[i].info.obj.lpParentID);
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.obj.lpParentID = \n", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.obj.ulObjType = 0x%08X",
					i,
					lpNotifications[i].info.obj.ulObjType);

				szPropNum =
					smartview::InterpretNumberAsStringProp(lpNotifications[i].info.obj.ulObjType, PR_OBJECT_TYPE);
				if (!szPropNum.empty())
				{
					Outputf(ulDbgLvl, fFile, false, L" = %ws", szPropNum.c_str());
				}

				Outputf(ulDbgLvl, fFile, false, L"\n");
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.obj.lpPropTagArray = \n", i);
				_OutputPropTagArray(ulDbgLvl, fFile, lpNotifications[i].info.obj.lpPropTagArray);
				break;
			case fnevIndexing:
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"lpNotifications[%u].info.ext.ulEvent = 0x%08X\n",
					i,
					lpNotifications[i].info.ext.ulEvent);
				sbin.cb = lpNotifications[i].info.ext.cb;
				sbin.lpb = lpNotifications[i].info.ext.pbEventParameters;
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.ext.pbEventParameters = \n", i);
				_OutputBinary(ulDbgLvl, fFile, sbin);
				if (INDEXING_SEARCH_OWNER == lpNotifications[i].info.ext.ulEvent &&
					sizeof(INDEX_SEARCH_PUSHER_PROCESS) == lpNotifications[i].info.ext.cb)
				{
					Outputf(
						ulDbgLvl, fFile, true, L"lpNotifications[%u].info.ext.ulEvent = INDEXING_SEARCH_OWNER\n", i);

					const auto lpidxExt =
						reinterpret_cast<INDEX_SEARCH_PUSHER_PROCESS*>(lpNotifications[i].info.ext.pbEventParameters);
					if (lpidxExt)
					{
						Outputf(ulDbgLvl, fFile, true, L"lpidxExt->dwPID = 0x%08X\n", lpidxExt->dwPID);
					}
				}

				break;
			}
		}
		Outputf(ulDbgLvl, fFile, true, L"End dumping notifications.\n");
	}

	void _OutputEntryList(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPENTRYLIST lpEntryList)
	{
		CHKPARAM;
		EARLYABORT;
		if (!lpEntryList) return;

		Outputf(ulDbgLvl, fFile, true, L"Dumping %u entry IDs.\n", lpEntryList->cValues);

		std::wstring szFlags;
		std::wstring szPropNum;

		for (ULONG i = 0; i < lpEntryList->cValues; i++)
		{
			Outputf(ulDbgLvl, fFile, true, L"lpEntryList->lpbin[%u]\n\t", i);
			_OutputBinary(ulDbgLvl, fFile, lpEntryList->lpbin[i]);
		}

		Outputf(ulDbgLvl, fFile, true, L"End dumping entry list.\n");
	}

	void _OutputProperty(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		_In_ LPSPropValue lpProp,
		_In_opt_ LPMAPIPROP lpObj,
		bool bRetryStreamProps)
	{
		CHKPARAM;
		EARLYABORT;

		if (!lpProp) return;

		LPSPropValue lpLargeProp = nullptr;
		const auto iIndent = 2;

		if (PROP_TYPE(lpProp->ulPropTag) == PT_ERROR && lpProp->Value.err == MAPI_E_NOT_ENOUGH_MEMORY && lpObj &&
			bRetryStreamProps)
		{
			lpLargeProp = mapi::GetLargeBinaryProp(lpObj, lpProp->ulPropTag);

			if (!lpLargeProp)
			{
				lpLargeProp = mapi::GetLargeStringProp(lpObj, lpProp->ulPropTag);
			}

			if (lpLargeProp && PT_ERROR != PROP_TYPE(lpLargeProp->ulPropTag))
			{
				lpProp = lpLargeProp;
			}
		}

		Outputf(
			ulDbgLvl,
			fFile,
			false,
			L"\t<property tag = \"0x%08X\" type = \"%ws\" >\n",
			lpProp->ulPropTag,
			proptype::TypeToString(lpProp->ulPropTag).c_str());

		auto propTagNames = proptags::PropTagToPropName(lpProp->ulPropTag, false);
		if (!propTagNames.bestGuess.empty())
			OutputXMLValue(
				ulDbgLvl,
				fFile,
				columns::PropXMLNames[columns::pcPROPBESTGUESS].uidName,
				propTagNames.bestGuess,
				false,
				iIndent);
		if (!propTagNames.otherMatches.empty())
			OutputXMLValue(
				ulDbgLvl,
				fFile,
				columns::PropXMLNames[columns::pcPROPOTHERNAMES].uidName,
				propTagNames.otherMatches,
				false,
				iIndent);

		auto namePropNames = cache::NameIDToStrings(lpProp->ulPropTag, lpObj, nullptr, nullptr, false);
		if (!namePropNames.guid.empty())
			OutputXMLValue(
				ulDbgLvl,
				fFile,
				columns::PropXMLNames[columns::pcPROPNAMEDIID].uidName,
				namePropNames.guid,
				false,
				iIndent);
		if (!namePropNames.name.empty())
			OutputXMLValue(
				ulDbgLvl,
				fFile,
				columns::PropXMLNames[columns::pcPROPNAMEDNAME].uidName,
				namePropNames.name,
				false,
				iIndent);

		auto prop = property::ParseProperty(lpProp);
		Output(ulDbgLvl, fFile, false, strings::StripCarriage(prop.toXML(iIndent)));

		auto szSmartView = smartview::InterpretPropSmartView(lpProp, lpObj, nullptr, nullptr, false, false);
		if (!szSmartView.empty())
		{
			OutputXMLValue(
				ulDbgLvl, fFile, columns::PropXMLNames[columns::pcPROPSMARTVIEW].uidName, szSmartView, true, iIndent);
		}

		Output(ulDbgLvl, fFile, false, L"\t</property>\n");

		if (lpLargeProp) MAPIFreeBuffer(lpLargeProp);
	}

	void _OutputProperties(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		ULONG cProps,
		_In_count_(cProps) LPSPropValue lpProps,
		_In_opt_ LPMAPIPROP lpObj,
		bool bRetryStreamProps)
	{
		CHKPARAM;
		EARLYABORT;

		if (cProps && !lpProps)
		{
			Output(ulDbgLvl, fFile, true, L"OutputProperties called with NULL lpProps!\n");
			return;
		}

		// Copy the list before we sort it or else we affect the caller
		// Don't worry about linked memory - we just need to sort the index
		const auto cbProps = cProps * sizeof(SPropValue);
		const auto lpSortedProps = mapi::allocate<LPSPropValue>(static_cast<ULONG>(cbProps));

		if (lpSortedProps)
		{
			memcpy(lpSortedProps, lpProps, cbProps);

			// sort the list first
			// insertion sort on lpSortedProps
			for (ULONG iUnsorted = 1; iUnsorted < cProps; iUnsorted++)
			{
				ULONG iLoc = 0;
				const auto NextItem = lpSortedProps[iUnsorted];
				for (iLoc = iUnsorted; iLoc > 0; iLoc--)
				{
					if (lpSortedProps[iLoc - 1].ulPropTag < NextItem.ulPropTag) break;
					lpSortedProps[iLoc] = lpSortedProps[iLoc - 1];
				}

				lpSortedProps[iLoc] = NextItem;
			}

			for (ULONG i = 0; i < cProps; i++)
			{
				_OutputProperty(ulDbgLvl, fFile, &lpSortedProps[i], lpObj, bRetryStreamProps);
			}
		}

		MAPIFreeBuffer(lpSortedProps);
	}

	void _OutputSRow(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ const _SRow* lpSRow, _In_opt_ LPMAPIPROP lpObj)
	{
		CHKPARAM;
		EARLYABORT;

		if (!lpSRow)
		{
			Output(ulDbgLvl, fFile, true, L"OutputSRow called with NULL lpSRow!\n");
			return;
		}

		if (lpSRow->cValues && !lpSRow->lpProps)
		{
			Output(ulDbgLvl, fFile, true, L"OutputSRow called with NULL lpSRow->lpProps!\n");
			return;
		}

		_OutputProperties(ulDbgLvl, fFile, lpSRow->cValues, lpSRow->lpProps, lpObj, false);
	}

	void _OutputSRowSet(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRowSet lpRowSet, _In_opt_ LPMAPIPROP lpObj)
	{
		CHKPARAM;
		EARLYABORT;

		if (!lpRowSet)
		{
			Output(ulDbgLvl, fFile, true, L"OutputSRowSet called with NULL lpRowSet!\n");
			return;
		}

		if (lpRowSet->cRows >= 1)
		{
			for (ULONG i = 0; i < lpRowSet->cRows; i++)
			{
				_OutputSRow(ulDbgLvl, fFile, &lpRowSet->aRow[i], lpObj);
			}
		}
	}

	void _OutputRestriction(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		_In_opt_ const _SRestriction* lpRes,
		_In_opt_ LPMAPIPROP lpObj)
	{
		CHKPARAM;
		EARLYABORT;

		if (!lpRes)
		{
			Output(ulDbgLvl, fFile, true, L"_OutputRestriction called with NULL lpRes!\n");
			return;
		}

		Output(ulDbgLvl, fFile, true, strings::StripCarriage(interpretprop::RestrictionToString(lpRes, lpObj)));
	}
} // namespace output