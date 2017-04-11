#include "stdafx.h"
#include <IO/MFCOutput.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/String.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/InterpretProp2.h>
#include <Interpret/SmartView/SmartView.h>
#include <MAPI/ColumnTags.h>
#include "Property/ParseProperty.h"
#include <UI/Dialogs/Editors/DbgView.h>

#ifdef CHECKFORMATPARAMS
#undef Outputf
#undef OutputToFilef
#undef DebugPrint
#undef DebugPrintEx
#endif

wstring g_szXMLHeader = L"<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>\n";
FILE* g_fDebugFile = nullptr;

void OpenDebugFile()
{
	// close out the old file handle if we have one
	CloseDebugFile();

	// only open the file if we really need to
	if (0 != RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD)
	{
		g_fDebugFile = MyOpenFile(RegKeys[regkeyDEBUG_FILE_NAME].szCurSTRING, false);
	}
}

void CloseDebugFile()
{
	if (g_fDebugFile) CloseFile(g_fDebugFile);
	g_fDebugFile = nullptr;
}

_Check_return_ ULONG GetDebugLevel()
{
	return RegKeys[regkeyDEBUG_TAG].ulCurDWORD;
}

void SetDebugLevel(ULONG ulDbgLvl)
{
	RegKeys[regkeyDEBUG_TAG].ulCurDWORD = ulDbgLvl;
}

// We've got our 'new' value here and also a debug output file name
// gonna set the new value
// gonna ensure our debug output file is open if we need it, closed if we don't
// gonna output some text if we just toggled on
void SetDebugOutputToFile(bool bDoOutput)
{
	// save our old value
	auto bOldDoOutput = 0 != RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD;

	// set the new value
	RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD = bDoOutput;

	// ensure we got a file if we need it
	OpenDebugFile();

	// output text if we just toggled on
	if (bDoOutput && !bOldDoOutput)
	{
		auto lpApp = AfxGetApp();

		if (lpApp)
		{
			DebugPrint(DBGGeneric, L"%ws: Debug printing to file enabled.\n", LPCTSTRToWstring(lpApp->m_pszAppName).c_str());
		}

		DebugPrintVersion(DBGVersionBanner);
	}
}

#define CHKPARAM ASSERT(DBGNoDebug != ulDbgLvl || fFile)

// quick check to see if we have anything to print - so we can avoid executing the call
#define EARLYABORT {if (!fFile && !RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD && !fIsSetv(ulDbgLvl)) return;}

_Check_return_ FILE* MyOpenFile(const wstring& szFileName, bool bNewFile)
{
	FILE* fOut = nullptr;
	auto szParams = L"a+"; // STRING_OK
	if (bNewFile) szParams = L"w"; // STRING_OK

	// _wfopen has been deprecated, but older compilers do not have _wfopen_s
	// Use the replacement when we're on VS 2005 or higher.
#if defined(_MSC_VER) && (_MSC_VER >= 1400)
	_wfopen_s(&fOut, szFileName.c_str(), szParams);
#else
	fOut = _wfopen(szFileName.c_str(), szParams);
#endif
	if (fOut)
	{
		return fOut;
	}

	// File IO failed - complain - not using error macros since we may not have debug output here
	auto dwErr = GetLastError();
	auto szSysErr = formatmessagesys(HRESULT_FROM_WIN32(dwErr));

	auto szErr = format(
		L"_tfopen failed, hRes = 0x%08X, dwErr = 0x%08X = \"%ws\"\n", // STRING_OK
		HRESULT_FROM_WIN32(dwErr),
		dwErr,
		szSysErr.c_str());

	OutputDebugStringW(szErr.c_str());
	return nullptr;
}

void CloseFile(_In_opt_ FILE* fFile)
{
	if (fFile) fclose(fFile);
}

void WriteFile(_In_ FILE* fFile, const wstring& szString)
{
	if (!szString.empty())
	{
		auto szStringA = wstringTostring(szString);
		fputs(szStringA.c_str(), fFile);
	}
}

void OutputThreadTime(ULONG ulDbgLvl)
{
	// Compute current time and thread for a time stamp
	wstring szThreadTime;

	SYSTEMTIME stLocalTime = {};
	FILETIME ftLocalTime = {};

	GetSystemTime(&stLocalTime);
	GetSystemTimeAsFileTime(&ftLocalTime);

	szThreadTime = format(
		L"0x%04x %02d:%02u:%02u.%03u%ws  %02u-%02u-%4u 0x%08X: ", // STRING_OK
		GetCurrentThreadId(),
		stLocalTime.wHour <= 12 ? stLocalTime.wHour : stLocalTime.wHour - 12,
		stLocalTime.wMinute,
		stLocalTime.wSecond,
		stLocalTime.wMilliseconds,
		stLocalTime.wHour <= 12 ? L"AM" : L"PM", // STRING_OK
		stLocalTime.wMonth,
		stLocalTime.wDay,
		stLocalTime.wYear,
		ulDbgLvl);
	OutputDebugStringW(szThreadTime.c_str());
#ifndef MRMAPI
	OutputToDbgView(szThreadTime);
#endif

	// print to to our debug output log file
	if (RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD && g_fDebugFile)
	{
		WriteFile(g_fDebugFile, szThreadTime);
	}
}

// The root of all debug output - call no debug output functions besides OutputDebugString from here!
void Output(ULONG ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, const wstring& szMsg)
{
	CHKPARAM;
	EARLYABORT;
	if (szMsg.empty()) return; // nothing to print? Cool!

	// print to debug output
	if (fIsSetv(ulDbgLvl))
	{
		if (bPrintThreadTime)
		{
			OutputThreadTime(ulDbgLvl);
		}

		OutputDebugStringW(szMsg.c_str());
#ifdef MRMAPI
		wprintf(L"%ws", szMsg.c_str());
#else
		OutputToDbgView(szMsg);
#endif

		// print to to our debug output log file
		if (RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD && g_fDebugFile)
		{
			WriteFile(g_fDebugFile, szMsg);
		}
	}

	// If we were given a file - send the output there
	if (fFile)
	{
		WriteFile(fFile, szMsg);
	}
}

void __cdecl Outputf(ULONG ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, LPCWSTR szMsg, ...)
{
	CHKPARAM;
	EARLYABORT;

	va_list argList = nullptr;
	va_start(argList, szMsg);
	Output(ulDbgLvl, fFile, bPrintThreadTime, formatV(szMsg, argList));
	va_end(argList);
}

void __cdecl OutputToFilef(_In_opt_ FILE* fFile, LPCWSTR szMsg, ...)
{
	if (!fFile) return;

	va_list argList = nullptr;
	va_start(argList, szMsg);
	if (argList)
	{
		Output(DBGNoDebug, fFile, true, formatV(szMsg, argList));
	}
	else
	{
		Output(DBGNoDebug, fFile, true, szMsg);
	}

	va_end(argList);

}

void __cdecl DebugPrint(ULONG ulDbgLvl, LPCWSTR szMsg, ...)
{
	if (!fIsSetv(ulDbgLvl) && !RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD) return;

	va_list argList = nullptr;
	va_start(argList, szMsg);
	if (argList)
	{
		Output(ulDbgLvl, nullptr, true, formatV(szMsg, argList));
	}
	else
	{
		Output(ulDbgLvl, nullptr, true, szMsg);
	}

	va_end(argList);
}

void __cdecl DebugPrintEx(ULONG ulDbgLvl, wstring& szClass, const wstring& szFunc, LPCWSTR szMsg, ...)
{
	if (!fIsSetv(ulDbgLvl) && !RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD) return;

	auto szMsgEx = format(L"%ws::%ws %ws", szClass.c_str(), szFunc.c_str(), szMsg); // STRING_OK
	va_list argList = nullptr;
	va_start(argList, szMsg);
	if (argList)
	{
		Output(ulDbgLvl, nullptr, true, formatV(szMsgEx.c_str(), argList));
	}
	else
	{
		Output(ulDbgLvl, nullptr, true, szMsgEx);
	}

	va_end(argList);
}

void OutputIndent(ULONG ulDbgLvl, _In_opt_ FILE* fFile, int iIndent)
{
	CHKPARAM;
	EARLYABORT;

	for (auto i = 0; i < iIndent; i++) Output(ulDbgLvl, fFile, false, L"\t");
}

void _OutputBinary(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ const SBinary& bin)
{
	CHKPARAM;
	EARLYABORT;

	Output(ulDbgLvl, fFile, false, BinToHexString(&bin, true));

	Output(ulDbgLvl, fFile, false, L"\n");
}

void _OutputNamedPropID(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPINAMEID lpName)
{
	CHKPARAM;
	EARLYABORT;

	if (!lpName) return;

	if (lpName->ulKind == MNID_ID)
	{
		Outputf(ulDbgLvl, fFile, true,
			L"\t\t: nmid ID: 0x%X\n", // STRING_OK
			lpName->Kind.lID);
	}
	else
	{
		Outputf(ulDbgLvl, fFile, true,
			L"\t\t: nmid Name: %ws\n", // STRING_OK
			lpName->Kind.lpwstrName);
	}

	Output(ulDbgLvl, fFile, false, GUIDToStringAndName(lpName->lpguid));
	Output(ulDbgLvl, fFile, false, L"\n");
}

void _OutputFormInfo(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMINFO lpMAPIFormInfo)
{
	CHKPARAM;
	EARLYABORT;
	if (!lpMAPIFormInfo) return;

	auto hRes = S_OK;

	LPSPropValue lpPropVals = nullptr;
	ULONG ulPropVals = NULL;
	LPMAPIVERBARRAY lpMAPIVerbArray = nullptr;
	LPMAPIFORMPROPARRAY lpMAPIFormPropArray = nullptr;

	Outputf(ulDbgLvl, fFile, true, L"Dumping verb and property set for form: %p\n", lpMAPIFormInfo);

	EC_H(GetPropsNULL(lpMAPIFormInfo, fMapiUnicode, &ulPropVals, &lpPropVals));
	if (lpPropVals)
	{
		_OutputProperties(ulDbgLvl, fFile, ulPropVals, lpPropVals, lpMAPIFormInfo, false);
		MAPIFreeBuffer(lpPropVals);
	}

	EC_MAPI(lpMAPIFormInfo->CalcVerbSet(NULL, &lpMAPIVerbArray)); // API doesn't support Unicode

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

	hRes = S_OK;
	EC_MAPI(lpMAPIFormInfo->CalcFormPropSet(NULL, &lpMAPIFormPropArray)); // API doesn't support Unicode

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
		Outputf(ulDbgLvl, fFile, true, L"\t\tProperty 0x%X\n",
			i);

		if (lpMAPIFormPropArray->aFormProp[i].ulFlags == MAPI_UNICODE)
		{
			Outputf(ulDbgLvl, fFile, true,
				L"\t\tProperty Name: %ws\n\t\tProperty Type: %ws\n\t\tSpecial Type: 0x%X\n\t\tNum Vals: 0x%X\n", // STRING_OK
				LPWSTR(lpMAPIFormPropArray->aFormProp[i].pszDisplayName),
				TypeToString(lpMAPIFormPropArray->aFormProp[i].nPropType).c_str(),
				lpMAPIFormPropArray->aFormProp[i].nSpecialType,
				lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable);
		}
		else
		{
			Outputf(ulDbgLvl, fFile, true,
				L"\t\tProperty Name: %hs\n\t\tProperty Type: %ws\n\t\tSpecial Type: 0x%X\n\t\tNum Vals: 0x%X\n", // STRING_OK
				LPSTR(lpMAPIFormPropArray->aFormProp[i].pszDisplayName),
				TypeToString(lpMAPIFormPropArray->aFormProp[i].nPropType).c_str(),
				lpMAPIFormPropArray->aFormProp[i].nSpecialType,
				lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable);
		}

		_OutputNamedPropID(ulDbgLvl, fFile, &lpMAPIFormPropArray->aFormProp[i].u.s1.nmidIdx);
		for (ULONG j = 0; j < lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable; j++)
		{
			if (lpMAPIFormPropArray->aFormProp[i].ulFlags == MAPI_UNICODE)
			{
				Outputf(ulDbgLvl, fFile, true,
					L"\t\t\tEnum 0x%X\nEnumVal Name: %ws\t\t\t\nEnumVal enumeration: 0x%X\n", // STRING_OK
					j,
					LPWSTR(lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].pszDisplayName),
					lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].nVal);
			}
			else
			{
				Outputf(ulDbgLvl, fFile, true,
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

	Outputf(ulDbgLvl,
		fFile,
		true,
		L"\tProp tag list, %u props\n", // STRING_OK
		lpTagsToDump->cValues);
	for (ULONG uCurProp = 0; uCurProp < lpTagsToDump->cValues; uCurProp++)
	{
		Outputf(ulDbgLvl,
			fFile,
			true,
			L"\t\tProp: %u = %ws\n", // STRING_OK
			uCurProp,
			TagToString(lpTagsToDump->aulPropTag[uCurProp], nullptr, false, true).c_str());
	}

	Output(ulDbgLvl, fFile, true, L"\tEnd Prop Tag List\n");
}

void _OutputTable(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPITABLE lpMAPITable)
{
	CHKPARAM;
	EARLYABORT;
	if (!lpMAPITable) return;

	auto hRes = S_OK;
	LPSRowSet lpRows = nullptr;

	EC_MAPI(lpMAPITable->SeekRow(
		BOOKMARK_BEGINNING,
		0,
		nullptr));

	Output(ulDbgLvl, fFile, false, g_szXMLHeader);
	Output(ulDbgLvl, fFile, false, L"<table>\n");

	for (;;)
	{
		hRes = S_OK;

		FreeProws(lpRows);
		lpRows = nullptr;
		EC_MAPI(lpMAPITable->QueryRows(
			20,
			NULL,
			&lpRows));
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

void _OutputNotifications(ULONG ulDbgLvl, _In_opt_ FILE* fFile, ULONG cNotify, _In_count_(cNotify) LPNOTIFICATION lpNotifications, _In_opt_ LPMAPIPROP lpObj)
{
	CHKPARAM;
	EARLYABORT;
	if (!lpNotifications) return;

	Outputf(ulDbgLvl, fFile, true, L"Dumping %u notifications.\n", cNotify);

	wstring szFlags;
	wstring szPropNum;

	for (ULONG i = 0; i < cNotify; i++)
	{
		Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].ulEventType = 0x%08X", i, lpNotifications[i].ulEventType);
		szFlags = InterpretFlags(flagNotifEventType, lpNotifications[i].ulEventType);
		if (!szFlags.empty())
		{
			Outputf(ulDbgLvl, fFile, false, L" = %ws", szFlags.c_str());
		}

		Output(ulDbgLvl, fFile, false, L"\n");

		SBinary sbin = { 0 };
		switch (lpNotifications[i].ulEventType)
		{
		case fnevCriticalError:
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.err.ulFlags = 0x%08X\n", i,
				lpNotifications[i].info.err.ulFlags);
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.err.scode = 0x%08X\n", i,
				lpNotifications[i].info.err.scode);
			sbin.cb = lpNotifications[i].info.err.cbEntryID;
			sbin.lpb = reinterpret_cast<LPBYTE>(lpNotifications[i].info.err.lpEntryID);
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.err.lpEntryID = ", i);
			_OutputBinary(ulDbgLvl, fFile, sbin);
			if (lpNotifications[i].info.err.lpMAPIError)
			{
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.err.lpMAPIError = %s\n", i,
					MAPIErrToString(
						lpNotifications[i].info.err.ulFlags,
						*lpNotifications[i].info.err.lpMAPIError).c_str());
			}

			break;
		case fnevExtended:
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.ext.ulEvent = 0x%08X\n", i,
				lpNotifications[i].info.ext.ulEvent);
			sbin.cb = lpNotifications[i].info.ext.cb;
			sbin.lpb = lpNotifications[i].info.ext.pbEventParameters;
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.ext.pbEventParameters = \n", i);
			_OutputBinary(ulDbgLvl, fFile, sbin);
			break;
		case fnevNewMail:
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.newmail.ulFlags = 0x%08X\n", i,
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
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.newmail.lpszMessageClass = \"%ws\"\n", i,
					reinterpret_cast<LPWSTR>(lpNotifications[i].info.newmail.lpszMessageClass));
			}
			else
			{
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.newmail.lpszMessageClass = \"%hs\"\n", i,
					LPSTR(lpNotifications[i].info.newmail.lpszMessageClass));
			}

			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.newmail.ulMessageFlags = 0x%08X", i,
				lpNotifications[i].info.newmail.ulMessageFlags);
			szPropNum = InterpretNumberAsStringProp(lpNotifications[i].info.newmail.ulMessageFlags, PR_MESSAGE_FLAGS);
			if (!szPropNum.empty())
			{
				Outputf(ulDbgLvl, fFile, false, L" = %ws", szPropNum.c_str());
			}

			Outputf(ulDbgLvl, fFile, false, L"\n");
			break;
		case fnevTableModified:
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.tab.ulTableEvent = 0x%08X", i,
				lpNotifications[i].info.tab.ulTableEvent);
			szFlags = InterpretFlags(flagTableEventType, lpNotifications[i].info.tab.ulTableEvent);
			if (!szFlags.empty())
			{
				Outputf(ulDbgLvl, fFile, false, L" = %ws", szFlags.c_str());
			}

			Outputf(ulDbgLvl, fFile, false, L"\n");

			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.tab.hResult = 0x%08X\n", i,
				lpNotifications[i].info.tab.hResult);
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.tab.propIndex = \n", i);
			_OutputProperty(ulDbgLvl, fFile,
				&lpNotifications[i].info.tab.propIndex,
				lpObj,
				false);
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.tab.propPrior = \n", i);
			_OutputProperty(ulDbgLvl, fFile,
				&lpNotifications[i].info.tab.propPrior,
				nullptr,
				false);
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.tab.row = \n", i);
			_OutputSRow(ulDbgLvl, fFile,
				&lpNotifications[i].info.tab.row,
				lpObj);
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
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.obj.ulObjType = 0x%08X", i,
				lpNotifications[i].info.obj.ulObjType);

			szPropNum = InterpretNumberAsStringProp(lpNotifications[i].info.obj.ulObjType, PR_OBJECT_TYPE);
			if (!szPropNum.empty())
			{
				Outputf(ulDbgLvl, fFile, false, L" = %ws", szPropNum.c_str());
			}

			Outputf(ulDbgLvl, fFile, false, L"\n");
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.obj.lpPropTagArray = \n", i);
			_OutputPropTagArray(ulDbgLvl, fFile,
				lpNotifications[i].info.obj.lpPropTagArray);
			break;
		case fnevIndexing:
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.ext.ulEvent = 0x%08X\n", i,
				lpNotifications[i].info.ext.ulEvent);
			sbin.cb = lpNotifications[i].info.ext.cb;
			sbin.lpb = lpNotifications[i].info.ext.pbEventParameters;
			Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.ext.pbEventParameters = \n", i);
			_OutputBinary(ulDbgLvl, fFile, sbin);
			if (INDEXING_SEARCH_OWNER == lpNotifications[i].info.ext.ulEvent &&
				sizeof(INDEX_SEARCH_PUSHER_PROCESS) == lpNotifications[i].info.ext.cb)
			{
				Outputf(ulDbgLvl, fFile, true, L"lpNotifications[%u].info.ext.ulEvent = INDEXING_SEARCH_OWNER\n", i);

				auto lpidxExt = reinterpret_cast<INDEX_SEARCH_PUSHER_PROCESS*>(lpNotifications[i].info.ext.pbEventParameters);
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

	wstring szFlags;
	wstring szPropNum;

	for (ULONG i = 0; i < lpEntryList->cValues; i++)
	{
		Outputf(ulDbgLvl, fFile, true, L"lpEntryList->lpbin[%u]\n\t", i);
		_OutputBinary(ulDbgLvl, fFile, lpEntryList->lpbin[i]);
	}

	Outputf(ulDbgLvl, fFile, true, L"End dumping entry list.\n");
}

void _OutputProperty(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropValue lpProp, _In_opt_ LPMAPIPROP lpObj, bool bRetryStreamProps)
{
	CHKPARAM;
	EARLYABORT;

	if (!lpProp) return;

	auto hRes = S_OK;
	LPSPropValue lpLargeProp = nullptr;
	auto iIndent = 2;

	if (PT_ERROR == PROP_TYPE(lpProp->ulPropTag) && MAPI_E_NOT_ENOUGH_MEMORY == lpProp->Value.err && lpObj && bRetryStreamProps)
	{
		WC_H(GetLargeBinaryProp(lpObj, lpProp->ulPropTag, &lpLargeProp));

		if (FAILED(hRes))
		{
			hRes = S_OK;
			WC_H(GetLargeStringProp(lpObj, lpProp->ulPropTag, &lpLargeProp));
		}

		if (SUCCEEDED(hRes) && lpLargeProp && PT_ERROR != PROP_TYPE(lpLargeProp->ulPropTag))
		{
			lpProp = lpLargeProp;
		}

		hRes = S_OK;
	}

	Outputf(ulDbgLvl, fFile, false, L"\t<property tag = \"0x%08X\" type = \"%ws\" >\n", lpProp->ulPropTag, TypeToString(lpProp->ulPropTag).c_str());

	auto propTagNames = PropTagToPropName(lpProp->ulPropTag, false);
	if (!propTagNames.bestGuess.empty()) OutputXMLValue(ulDbgLvl, fFile, PropXMLNames[pcPROPBESTGUESS].uidName, propTagNames.bestGuess, false, iIndent);
	if (!propTagNames.otherMatches.empty()) OutputXMLValue(ulDbgLvl, fFile, PropXMLNames[pcPROPOTHERNAMES].uidName, propTagNames.otherMatches, false, iIndent);

	auto namePropNames = NameIDToStrings(
		lpProp->ulPropTag,
		lpObj,
		nullptr,
		nullptr,
		false);
	if (!namePropNames.guid.empty()) OutputXMLValue(ulDbgLvl, fFile, PropXMLNames[pcPROPNAMEDIID].uidName, namePropNames.guid, false, iIndent);
	if (!namePropNames.name.empty()) OutputXMLValue(ulDbgLvl, fFile, PropXMLNames[pcPROPNAMEDNAME].uidName, namePropNames.name, false, iIndent);

	auto prop = ParseProperty(lpProp);
	Output(ulDbgLvl, fFile, false, StripCarriage(prop.toXML(iIndent)));

	auto szSmartView = InterpretPropSmartView(
		lpProp,
		lpObj,
		nullptr,
		nullptr,
		false,
		false);
	if (!szSmartView.empty())
	{
		OutputXMLValue(ulDbgLvl, fFile, PropXMLNames[pcPROPSMARTVIEW].uidName, szSmartView, true, iIndent);
	}

	Output(ulDbgLvl, fFile, false, L"\t</property>\n");

	if (lpLargeProp) MAPIFreeBuffer(lpLargeProp);
}

void _OutputProperties(ULONG ulDbgLvl, _In_opt_ FILE* fFile, ULONG cProps, _In_count_(cProps) LPSPropValue lpProps, _In_opt_ LPMAPIPROP lpObj, bool bRetryStreamProps)
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
	LPSPropValue lpSortedProps = nullptr;
	size_t cbProps = cProps * sizeof(SPropValue);
	MAPIAllocateBuffer(static_cast<ULONG>(cbProps), reinterpret_cast<LPVOID*>(&lpSortedProps));

	if (lpSortedProps)
	{
		memcpy(lpSortedProps, lpProps, cbProps);

		// sort the list first
		// insertion sort on lpSortedProps
		for (ULONG iUnsorted = 1; iUnsorted < cProps; iUnsorted++)
		{
			ULONG iLoc = 0;
			auto NextItem = lpSortedProps[iUnsorted];
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

void _OutputSRow(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRow lpSRow, _In_opt_ LPMAPIPROP lpObj)
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

void _OutputRestriction(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_opt_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj)
{
	CHKPARAM;
	EARLYABORT;

	if (!lpRes)
	{
		Output(ulDbgLvl, fFile, true, L"_OutputRestriction called with NULL lpRes!\n");
		return;
	}

	Output(ulDbgLvl, fFile, true, StripCarriage(RestrictionToString(lpRes, lpObj)));
}

#define MAXBYTES 4096
void _OutputStream(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSTREAM lpStream)
{
	CHKPARAM;
	EARLYABORT;
	auto hRes = S_OK;
	BYTE bBuf[MAXBYTES + 2]; // Allocate some extra for NULL terminators - 2 for Unicode
	ULONG ulNumBytes = 0;
	LARGE_INTEGER li = { 0 };

	if (!lpStream)
	{
		Output(ulDbgLvl, fFile, true, L"OutputStream called with NULL lpStream!\n");
		return;
	}

	WC_H_MSG(lpStream->Seek(
		li,
		STREAM_SEEK_SET,
		nullptr), IDS_STREAMSEEKFAILED);

	if (S_OK == hRes) do
	{
		hRes = S_OK;
		ulNumBytes = 0;
		EC_MAPI(lpStream->Read(
			bBuf,
			MAXBYTES,
			&ulNumBytes));

		if (ulNumBytes > 0)
		{
			bBuf[ulNumBytes] = 0;
			bBuf[ulNumBytes + 1] = 0; // In case we are in Unicode
			// TODO: Check how this works in Unicode vs ANSI
			Output(ulDbgLvl, fFile, true, StripCarriage(LPCTSTRToWstring(reinterpret_cast<TCHAR*>(bBuf))));
		}
	} while (ulNumBytes > 0);
}

void _OutputVersion(ULONG ulDbgLvl, _In_opt_ FILE* fFile)
{
	CHKPARAM;
	EARLYABORT;
	wchar_t szFullPath[MAX_PATH];
	auto hRes = S_OK;
	DWORD dwRet = 0;

	// Get version information from the application.
	EC_D(dwRet, GetModuleFileNameW(nullptr, szFullPath, _countof(szFullPath)));

	if (S_OK == hRes)
	{
		DWORD dwVerInfoSize = 0;

		EC_D(dwVerInfoSize, GetFileVersionInfoSizeW(szFullPath, nullptr));

		if (dwVerInfoSize)
		{
			// If we were able to get the information, process it.
			auto pbData = new BYTE[dwVerInfoSize];
			if (pbData == nullptr) return;

			BOOL bRet = false;
			EC_D(bRet, GetFileVersionInfoW(
				szFullPath,
				NULL,
				dwVerInfoSize,
				static_cast<void*>(pbData)));

			if (pbData)
			{
				struct LANGANDCODEPAGE {
					WORD wLanguage;
					WORD wCodePage;
				} *lpTranslate = { nullptr };

				UINT cbTranslate = 0;

				// Read the list of languages and code pages.
				EC_B(VerQueryValueW(
					pbData,
					L"\\VarFileInfo\\Translation", // STRING_OK
					reinterpret_cast<LPVOID*>(&lpTranslate),
					&cbTranslate));

				// Read the file description for each language and code page.

				if (S_OK == hRes && lpTranslate)
				{
					for (UINT iCodePages = 0; iCodePages < cbTranslate / sizeof(LANGANDCODEPAGE); iCodePages++)
					{
						auto szSubBlock = format(
							L"\\StringFileInfo\\%04x%04x\\", // STRING_OK
							lpTranslate[iCodePages].wLanguage,
							lpTranslate[iCodePages].wCodePage);

						// Load all our strings
						for (auto iVerString = IDS_VER_FIRST; iVerString <= IDS_VER_LAST; iVerString++)
						{
							UINT cchVer = 0;
							wchar_t* lpszVer = nullptr;
							auto szVerString = loadstring(iVerString);
							auto szQueryString = szSubBlock + szVerString;
							hRes = S_OK;

							EC_B(VerQueryValueW(
								static_cast<void*>(pbData),
								szQueryString.c_str(),
								reinterpret_cast<void**>(&lpszVer),
								&cchVer));

							if (S_OK == hRes && cchVer && lpszVer)
							{
								Outputf(ulDbgLvl, fFile, true, L"%ws: %ws\n", szVerString.c_str(), lpszVer);
							}
						}
					}
				}
			}

			delete[] pbData;
		}
	}
}

void OutputCDataOpen(ULONG ulDbgLvl, _In_opt_ FILE* fFile)
{
	Output(ulDbgLvl, fFile, false, L"<![CDATA[");
}

void OutputCDataClose(ULONG ulDbgLvl, _In_opt_ FILE* fFile)
{
	Output(ulDbgLvl, fFile, false, L"]]>");
}

void OutputXMLValue(ULONG ulDbgLvl, _In_opt_ FILE* fFile, UINT uidTag, const wstring& szValue, bool bWrapCData, int iIndent)
{
	CHKPARAM;
	EARLYABORT;
	if (szValue.empty() || !uidTag) return;

	if (!szValue[0]) return;

	auto szTag = loadstring(uidTag);

	OutputIndent(ulDbgLvl, fFile, iIndent);
	Outputf(ulDbgLvl, fFile, false, L"<%ws>", szTag.c_str());
	if (bWrapCData)
	{
		OutputCDataOpen(ulDbgLvl, fFile);
	}

	Output(ulDbgLvl, fFile, false, StripCarriage(szValue));

	if (bWrapCData)
	{
		OutputCDataClose(ulDbgLvl, fFile);
	}

	Outputf(ulDbgLvl, fFile, false, L"</%ws>\n", szTag.c_str());
}