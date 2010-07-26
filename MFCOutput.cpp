// MFCOutput.cpp

#include "stdafx.h"
#include "MFCOutput.h"
#include "MapiFunctions.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "PropTagArray.h"
#include "DbgView.h"

FILE* g_fDebugFile = NULL;

void OpenDebugFile()
{
	// close out the old file handle if we have one
	CloseDebugFile();

	// only open the file if we really need to
	if (RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD == TRUE)
	{
#ifdef UNICODE
		g_fDebugFile = OpenFile(RegKeys[regkeyDEBUG_FILE_NAME].szCurSTRING,false);
#else
		LPWSTR szDebugFileW = NULL;
		(void) AnsiToUnicode(RegKeys[regkeyDEBUG_FILE_NAME].szCurSTRING,&szDebugFileW);
		g_fDebugFile = OpenFile(szDebugFileW,false);
		delete[] szDebugFileW;
#endif
	}
} // OpenDebugFile

void CloseDebugFile()
{
	if (g_fDebugFile) CloseFile(g_fDebugFile);
	g_fDebugFile = NULL;
} // CloseDebugFile

_Check_return_ ULONG GetDebugLevel()
{
	return RegKeys[regkeyDEBUG_TAG].ulCurDWORD;
} // GetDebugLevel

void SetDebugLevel(ULONG ulDbgLvl)
{
	RegKeys[regkeyDEBUG_TAG].ulCurDWORD = ulDbgLvl;
} // SetDebugLevel

// We've got our 'new' value here and also a debug output file name
// gonna set the new value
// gonna ensure our debug output file is open if we need it, closed if we don't
// gonna output some text if we just toggled on
void SetDebugOutputToFile(BOOL bDoOutput)
{
	// save our old value
	BOOL bOldDoOutput = RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD;

	// set the new value
	RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD = bDoOutput;

	// ensure we got a file if we need it
	OpenDebugFile();

	// output text if we just toggled on
	if (bDoOutput && !bOldDoOutput)
	{
		CWinApp* lpApp = NULL;

		lpApp = AfxGetApp();

		if (lpApp)
		{
			DebugPrint(DBGGeneric,_T("%s: Debug printing to file enabled.\n"),lpApp->m_pszAppName);
		}
		DebugPrintVersion(DBGVersionBanner);
	}
} // SetDebugOutputToFile

#define CHKPARAM ASSERT(DBGNoDebug != ulDbgLvl || fFile)

// quick check to see if we have anything to print - so we can avoid executing the call
#define EARLYABORT {if (!fFile && !RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD && !fIsSet(ulDbgLvl)) return;}

_Check_return_ FILE* OpenFile(_In_z_ LPCWSTR szFileName, BOOL bNewFile)
{
	static TCHAR szErr[256]; // buffer for catastrophic failures
	FILE* fOut = NULL;
	LPCWSTR szParams = L"a+"; // STRING_OK
	if (bNewFile) szParams = L"w"; // STRING_OK

	// _wfopen has been deprecated, but older compilers do not have _wfopen_s
	// Use the replacement when we're on VS 2005 or higher.
#if defined(_MSC_VER) && (_MSC_VER >= 1400)
	_wfopen_s(&fOut,szFileName,szParams);
#else
	fOut = _wfopen(szFileName,szParams);
#endif
	if (fOut)
	{
		return fOut;
	}
	else
	{
		// File IO failed - complain - not using error macros since we may not have debug output here
		DWORD dwErr = GetLastError();
		HRESULT hRes = HRESULT_FROM_WIN32(dwErr);
		LPTSTR szSysErr = NULL;
		FormatMessage(
			FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
			0,
			dwErr,
			0,
			(LPTSTR)&szSysErr,
			0,
			0);

		hRes = StringCchPrintf(szErr,_countof(szErr),
			_T("_tfopen failed, hRes = 0x%08X, dwErr = 0x%08X = \"%s\"\n"), // STRING_OK
			HRESULT_FROM_WIN32(dwErr),
			dwErr,
			szSysErr);

		OutputDebugString(szErr);
		LocalFree(szSysErr);
		return NULL;
	}
} // OpenFile

void CloseFile(_In_opt_ FILE* fFile)
{
	if (fFile) fclose(fFile);
} // CloseFile

// The root of all debug output - call no debug output functions besides OutputDebugString from here!
void _Output(ULONG ulDbgLvl, _In_opt_ FILE* fFile, BOOL bPrintThreadTime, _In_opt_z_ LPCTSTR szMsg)
{
	CHKPARAM;
	EARLYABORT;
	static TCHAR szErr[64]; // buffer for catastrophic failures

	HRESULT hRes = S_OK;

	if (!szMsg) return; // nothing to print? Cool!

	// print to debug output
	if (fIsSet(ulDbgLvl))
	{
		// Compute current time and thread for a time stamp
		TCHAR		szThreadTime[MAX_PATH];

		if (bPrintThreadTime)
		{
			SYSTEMTIME	stLocalTime;
			FILETIME	ftLocalTime;
			GetSystemTime(&stLocalTime);
			GetSystemTimeAsFileTime(&ftLocalTime);

			EC_H(StringCchPrintf(szThreadTime,
				_countof(szThreadTime),
				_T("0x%04x %02d:%02d:%02d.%03d%s  %02d-%02d-%4d 0x%08X: "), // STRING_OK
				GetCurrentThreadId(),
				(stLocalTime.wHour <= 12)?stLocalTime.wHour:stLocalTime.wHour-12,
				stLocalTime.wMinute,
				stLocalTime.wSecond,
				stLocalTime.wMilliseconds,
				(stLocalTime.wHour <= 12)?_T("AM"):_T("PM"), // STRING_OK
				stLocalTime.wMonth,
				stLocalTime.wDay,
				stLocalTime.wYear,
				ulDbgLvl));
			OutputDebugString(szThreadTime);
			OutputToDbgView(szThreadTime);
		}

		OutputDebugString(szMsg);
		OutputToDbgView(szMsg);

		// print to to our debug output log file
		if (RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD && g_fDebugFile)
		{
			if (bPrintThreadTime) _fputts(szThreadTime,g_fDebugFile);
			_fputts(szMsg,g_fDebugFile);
		}
	}

	// If we were given a file - send the output there
	if (fFile)
	{
		_fputts(szMsg,fFile);
	}
} // _Output

void __cdecl Outputf(ULONG ulDbgLvl, _In_opt_ FILE* fFile, BOOL bPrintThreadTime, _Printf_format_string_ LPCTSTR szMsg, ...)
{
	CHKPARAM;
	EARLYABORT;
	HRESULT hRes = S_OK;

	if (!szMsg)
	{
		_Output(ulDbgLvl, fFile, true, _T("Output called with NULL szMsg!\n"));
		return;
	}

	va_list argList = NULL;
	va_start(argList, szMsg);

	TCHAR szDebugString[4096];
	WC_H(StringCchVPrintf(szDebugString, _countof(szDebugString), szMsg, argList));
	if (FAILED(hRes))
	{
		_Output(DBGFatalError,NULL, true,_T("Debug output string not large enough to print everything to it\n"));
		// Since this function was 'safe', we've still got something we can print - send it on.
	}
	va_end(argList);

	_Output(ulDbgLvl, fFile, bPrintThreadTime, szDebugString);
} // Outputf

void __cdecl OutputToFilef(_In_opt_ FILE* fFile, _Printf_format_string_ LPCTSTR szMsg, ...)
{
	HRESULT hRes = S_OK;

	if (!fFile) return;

	if (!szMsg)
	{
		_Output(DBGFatalError, fFile, true, _T("OutputToFilef called with NULL szMsg!\n"));
		return;
	}

	va_list argList = NULL;
	va_start(argList, szMsg);

	TCHAR szDebugString[4096];
	hRes = StringCchVPrintf(szDebugString, _countof(szDebugString), szMsg, argList);
	if (S_OK != hRes)
	{
		_Output(DBGFatalError,NULL, true,_T("Debug output string not large enough to print everything to it\n"));
		// Since this function was 'safe', we've still got something we can print - send it on.
	}
	va_end(argList);

	_Output(DBGNoDebug, fFile, false, szDebugString);
} // OutputToFilef

void __cdecl DebugPrint(ULONG ulDbgLvl, _Printf_format_string_ LPCTSTR szMsg,...)
{
	HRESULT hRes = S_OK;

	if (!fIsSet(ulDbgLvl) && !RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD) return;

	va_list argList = NULL;
	va_start(argList, szMsg);

	if (argList)
	{
		TCHAR szDebugString[4096];
		hRes = StringCchVPrintf(szDebugString, _countof(szDebugString), szMsg, argList);
		if (hRes == S_OK)
			_Output(ulDbgLvl,NULL, true, szDebugString);
	}
	else
		_Output(ulDbgLvl,NULL, true, szMsg);
	va_end(argList);
} // DebugPrint

void __cdecl DebugPrintEx(ULONG ulDbgLvl, _In_z_ LPCTSTR szClass, _In_z_ LPCTSTR szFunc, _Printf_format_string_ LPCTSTR szMsg, ...)
{
	HRESULT hRes = S_OK;
	static TCHAR szMsgEx[1024];

	if (!fIsSet(ulDbgLvl) && !RegKeys[regkeyDEBUG_TO_FILE].ulCurDWORD) return;

	hRes = StringCchPrintf(szMsgEx,_countof(szMsgEx),_T("%s::%s %s"),szClass, szFunc, szMsg); // STRING_OK
	if (hRes == S_OK)
	{
		va_list argList = NULL;
		va_start(argList, szMsg);

		TCHAR szDebugString[4096];
		hRes = StringCchVPrintf(szDebugString, _countof(szDebugString), szMsgEx, argList);
		va_end(argList);
		if (hRes == S_OK)
			_Output(ulDbgLvl, NULL, true, szDebugString);
	}
} // DebugPrintEx

void _OutputBinary(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSBinary lpBin)
{
	CHKPARAM;
	EARLYABORT;

	if (!lpBin)
	{
		_Output(DBGFatalError,fFile, true,_T("OutputBinary called with NULL lpBin!\n"));
		return;
	}

	_Output(ulDbgLvl, fFile, false, BinToHexString(lpBin,true));

	_Output(ulDbgLvl, fFile,false,_T("\n"));
} // _OutputBinary

void _OutputNamedPropID(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPINAMEID lpName)
{
	CHKPARAM;
	EARLYABORT;

	if (!lpName) return;

	if (lpName->ulKind == MNID_ID)
	{
		Outputf(ulDbgLvl,fFile,true,
			_T("\t\t: nmid ID: 0x%X\n"), // STRING_OK
			lpName->Kind.lID);
	}
	else
	{
		Outputf(ulDbgLvl,fFile,true,
			_T("\t\t: nmid Name: %ws\n"), // STRING_OK
			lpName->Kind.lpwstrName);
	}
	LPTSTR szGuid = GUIDToStringAndName(lpName->lpguid);
	if (szGuid)
	{
		_Output(ulDbgLvl, fFile, false, szGuid);
		delete[] szGuid;
		Outputf(ulDbgLvl,fFile,false,_T("\n"));
	}
} // _OutputNamedPropID

void _OutputFormInfo(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMINFO lpMAPIFormInfo)
{
	CHKPARAM;
	EARLYABORT;
	if (!lpMAPIFormInfo) return;

	HRESULT hRes = S_OK;

	LPSPropValue lpPropVals = NULL;
	ULONG ulPropVals = NULL;
	LPMAPIVERBARRAY lpMAPIVerbArray = NULL;
	LPMAPIFORMPROPARRAY lpMAPIFormPropArray = NULL;

	Outputf(ulDbgLvl,fFile, true, _T("Dumping verb and property set for form: %p\n"), lpMAPIFormInfo);

	EC_H(GetPropsNULL(lpMAPIFormInfo,fMapiUnicode,&ulPropVals,&lpPropVals));
	if (lpPropVals)
	{
		_OutputProperties(ulDbgLvl,fFile,ulPropVals,lpPropVals,lpMAPIFormInfo);
		MAPIFreeBuffer(lpPropVals);
	}

	EC_H(lpMAPIFormInfo->CalcVerbSet(NULL, &lpMAPIVerbArray)); // API doesn't support Unicode

	if (lpMAPIVerbArray)
	{
		Outputf(ulDbgLvl,fFile, true, _T("\t0x%X verbs:\n"), lpMAPIVerbArray->cMAPIVerb);
		for (ULONG i = 0 ; i < lpMAPIVerbArray->cMAPIVerb ; i++)
		{
			Outputf(
				ulDbgLvl,
				fFile,
				true,
				_T("\t\tVerb 0x%X\n"), // STRING_OK
				i);
			if (lpMAPIVerbArray->aMAPIVerb[i].ulFlags == MAPI_UNICODE)
			{
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					_T("\t\tDoVerb value: 0x%X\n\t\tUnicode Name: %ws\n\t\tFlags: 0x%X\n\t\tAttributes: 0x%X\n"), // STRING_OK
					lpMAPIVerbArray->aMAPIVerb[i].lVerb,
					(LPWSTR) lpMAPIVerbArray->aMAPIVerb[i].szVerbname,
					lpMAPIVerbArray->aMAPIVerb[i].fuFlags,
					lpMAPIVerbArray->aMAPIVerb[i].grfAttribs);
			}
			else
			{
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					_T("\t\tDoVerb value: 0x%X\n\t\tANSI Name: %hs\n\t\tFlags: 0x%X\n\t\tAttributes: 0x%X\n"), // STRING_OK
					lpMAPIVerbArray->aMAPIVerb[i].lVerb,
					(LPSTR) lpMAPIVerbArray->aMAPIVerb[i].szVerbname,
					lpMAPIVerbArray->aMAPIVerb[i].fuFlags,
					lpMAPIVerbArray->aMAPIVerb[i].grfAttribs);
			}
		}
		MAPIFreeBuffer(lpMAPIVerbArray);
	}

	hRes = S_OK;
	EC_H(lpMAPIFormInfo->CalcFormPropSet(NULL, &lpMAPIFormPropArray)); // API doesn't support Unicode

	if (lpMAPIFormPropArray)
	{
		_OutputFormPropArray(ulDbgLvl,fFile,lpMAPIFormPropArray);
		MAPIFreeBuffer(lpMAPIFormPropArray);
	}
} // _OutputFormInfo

void _OutputFormPropArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMPROPARRAY lpMAPIFormPropArray)
{
	Outputf(ulDbgLvl,fFile, true, _T("\t0x%X Properties:\n"), lpMAPIFormPropArray->cProps);
	for (ULONG i = 0 ; i < lpMAPIFormPropArray->cProps ; i++)
	{
		Outputf(ulDbgLvl,fFile, true, _T("\t\tProperty 0x%X\n"),
			i);

		if (lpMAPIFormPropArray->aFormProp[i].ulFlags == MAPI_UNICODE)
		{
			Outputf(ulDbgLvl,fFile, true,
				_T("\t\tProperty Name: %ws\n\t\tProperty Type: %s\n\t\tSpecial Type: 0x%X\n\t\tNum Vals: 0x%X\n"), // STRING_OK
				(LPWSTR) lpMAPIFormPropArray->aFormProp[i].pszDisplayName,
				(LPCTSTR) TypeToString(lpMAPIFormPropArray->aFormProp[i].nPropType),
				lpMAPIFormPropArray->aFormProp[i].nSpecialType,
				lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable);
		}
		else
		{
			Outputf(ulDbgLvl,fFile, true,
				_T("\t\tProperty Name: %hs\n\t\tProperty Type: %s\n\t\tSpecial Type: 0x%X\n\t\tNum Vals: 0x%X\n"), // STRING_OK
				(LPSTR) lpMAPIFormPropArray->aFormProp[i].pszDisplayName,
				(LPCTSTR) TypeToString(lpMAPIFormPropArray->aFormProp[i].nPropType),
				lpMAPIFormPropArray->aFormProp[i].nSpecialType,
				lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable);
		}
		_OutputNamedPropID(ulDbgLvl,fFile,&lpMAPIFormPropArray->aFormProp[i].u.s1.nmidIdx);
		for (ULONG j = 0; j < lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable ; j++)
		{
			if (lpMAPIFormPropArray->aFormProp[i].ulFlags == MAPI_UNICODE)
			{
				Outputf(ulDbgLvl,fFile, true,
					_T("\t\t\tEnum 0x%X\nEnumVal Name: %ws\t\t\t\nEnumVal enumeration: 0x%X\n"), // STRING_OK
					j,
					(LPWSTR) lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].pszDisplayName,
					lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].nVal);
			}
			else
			{
				Outputf(ulDbgLvl,fFile, true,
					_T("\t\t\tEnum 0x%X\nEnumVal Name: %hs\t\t\t\nEnumVal enumeration: 0x%X\n"), // STRING_OK
					j,
					(LPSTR) lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].pszDisplayName,
					lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].nVal);
			}
		}
		_OutputNamedPropID(ulDbgLvl,fFile,&lpMAPIFormPropArray->aFormProp[i].nmid);
	}
} // _OutputFormPropArray

void _OutputPropTagArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropTagArray lpTagsToDump)
{
	CHKPARAM;
	EARLYABORT;
	if (!lpTagsToDump) return;

	Outputf(ulDbgLvl,
		fFile,
		true,
		_T("\tProp tag list, %d props\n"), // STRING_OK
		lpTagsToDump->cValues);
	ULONG uCurProp = 0;
	for (uCurProp = 0;uCurProp < lpTagsToDump->cValues; uCurProp++)
	{
		Outputf(ulDbgLvl,
			fFile,
			true,
			_T("\t\tProp: %d = %s\n"), // STRING_OK
			uCurProp,
			(LPCTSTR) TagToString(lpTagsToDump->aulPropTag[uCurProp],NULL,false,true));
	}
	_Output(ulDbgLvl,fFile, true,_T("\tEnd Prop Tag List\n"));
} // _OutputPropTagArray

void _OutputTable(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPITABLE lpMAPITable)
{
	CHKPARAM;
	EARLYABORT;
	if (!lpMAPITable) return;

	HRESULT			hRes = S_OK;
	LPSRowSet		lpRows = NULL;

	EC_H(lpMAPITable->SeekRow(
		BOOKMARK_BEGINNING,
		0,
		NULL));
	hRes = S_OK; // don't let failure here fail the whole op

	for (;;)
	{
		hRes = S_OK;

		MAPIFreeBuffer(lpRows);
		lpRows = NULL;
		EC_H(lpMAPITable->QueryRows(
			20,
			NULL,
			&lpRows));
		if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

		ULONG iCurRow = 0;
		for (iCurRow = 0; iCurRow<lpRows->cRows; iCurRow++)
		{
			hRes = S_OK;
			Outputf(ulDbgLvl,fFile,true,_T("Row 0x%X of 0x%X\n"),iCurRow+1,lpRows->cRows);
			_OutputSRow(ulDbgLvl,fFile,&lpRows->aRow[iCurRow],NULL);
		}
	}

	MAPIFreeBuffer(lpRows);
	lpRows = NULL;
} // _OutputTable

void _OutputNotifications(ULONG ulDbgLvl, _In_opt_ FILE* fFile, ULONG cNotify, _In_count_(cNotify) LPNOTIFICATION lpNotifications)
{
	CHKPARAM;
	EARLYABORT;
	if (!lpNotifications) return;

	Outputf(ulDbgLvl,fFile,true,_T("Dumping %d notifications.\n"), cNotify);

	LPTSTR szFlags = NULL;

	for (ULONG i=0; i<cNotify; i++)
	{
		Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].ulEventType = 0x%08X"),i,lpNotifications[i].ulEventType);
		InterpretFlags(flagNotifEventType, lpNotifications[i].ulEventType, &szFlags);
		if (szFlags)
		{
			Outputf(ulDbgLvl,fFile,false,_T(" = %s"),szFlags);
		}
		delete[] szFlags;
		szFlags = NULL;
		Outputf(ulDbgLvl,fFile,false,_T("\n"));

		switch (lpNotifications[i].ulEventType)
		{
		case(fnevCriticalError) :
			{
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.err.ulFlags = 0x%08X\n"),i,
					lpNotifications[i].info.err.ulFlags);
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.err.scode = 0x%08X\n"),i,
					lpNotifications[i].info.err.scode);
				SBinary sbin = {0};
				sbin.cb = lpNotifications[i].info.err.cbEntryID;
				sbin.lpb = (LPBYTE) lpNotifications[i].info.err.lpEntryID;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.err.lpEntryID = "),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.err.lpMAPIError = %s\n"),i,
					(LPCTSTR) MAPIErrToString(
					lpNotifications[i].info.err.ulFlags,
					lpNotifications[i].info.err.lpMAPIError));
			}
			break;
		case(fnevExtended) :
			{
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.ext.ulEvent = 0x%08X\n"),i,
					lpNotifications[i].info.ext.ulEvent);
				SBinary sbin = {0};
				sbin.cb = lpNotifications[i].info.ext.cb;
				sbin.lpb = lpNotifications[i].info.ext.pbEventParameters;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.ext.pbEventParameters = \n"),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);
			}
			break;
		case(fnevNewMail) :
			{
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.newmail.ulFlags = 0x%08X\n"),i,
					lpNotifications[i].info.newmail.ulFlags);
				SBinary sbin = {0};
				sbin.cb = lpNotifications[i].info.newmail.cbEntryID;
				sbin.lpb = (LPBYTE) lpNotifications[i].info.newmail.lpEntryID;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.newmail.lpEntryID = \n"),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);
				sbin.cb = lpNotifications[i].info.newmail.cbParentID;
				sbin.lpb = (LPBYTE) lpNotifications[i].info.newmail.lpParentID;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.newmail.lpParentID = \n"),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);

				if (lpNotifications[i].info.newmail.ulFlags & MAPI_UNICODE)
				{
					Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.newmail.lpszMessageClass = \"%ws\"\n"),i,
						(LPWSTR) lpNotifications[i].info.newmail.lpszMessageClass);
				}
				else
				{
					Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.newmail.lpszMessageClass = \"%hs\"\n"),i,
						(LPSTR) lpNotifications[i].info.newmail.lpszMessageClass);
				}

				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.newmail.ulMessageFlags = 0x%08X"),i,
					lpNotifications[i].info.newmail.ulMessageFlags);
				InterpretFlags(PROP_ID(PR_MESSAGE_FLAGS),lpNotifications[i].info.newmail.ulMessageFlags,&szFlags);
				if (szFlags)
				{
					Outputf(ulDbgLvl,fFile,false,_T(" = %s"),szFlags);
				}
				delete[] szFlags;
				szFlags = NULL;
				Outputf(ulDbgLvl,fFile,false,_T("\n"));
			}
			break;
		case(fnevTableModified) :
			{
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.tab.ulTableEvent = 0x%08X"),i,
					lpNotifications[i].info.tab.ulTableEvent);
				InterpretFlags(flagTableEventType,lpNotifications[i].info.tab.ulTableEvent, &szFlags);
				if (szFlags)
				{
					Outputf(ulDbgLvl,fFile,false,_T(" = %s"),szFlags);
				}
				delete[] szFlags;
				szFlags = NULL;
				Outputf(ulDbgLvl,fFile,false,_T("\n"));

				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.tab.hResult = 0x%08X\n"),i,
					lpNotifications[i].info.tab.hResult);
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.tab.propIndex = \n"),i);
				_OutputProperty(ulDbgLvl,fFile,
					&lpNotifications[i].info.tab.propIndex,
					NULL);
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.tab.propPrior = \n"),i);
				_OutputProperty(ulDbgLvl,fFile,
					&lpNotifications[i].info.tab.propPrior,
					NULL);
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.tab.row = \n"),i);
				_OutputSRow(ulDbgLvl,fFile,
					&lpNotifications[i].info.tab.row,
					NULL);
			}
			break;
		case(fnevObjectCopied) :
		case(fnevObjectCreated) :
		case(fnevObjectDeleted) :
		case(fnevObjectModified) :
		case(fnevObjectMoved) :
		case(fnevSearchComplete) :
			{
				SBinary sbin = {0};
				sbin.cb = lpNotifications[i].info.obj.cbOldID;
				sbin.lpb = (LPBYTE) lpNotifications[i].info.obj.lpOldID;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.obj.lpOldID = \n"),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);
				sbin.cb = lpNotifications[i].info.obj.cbOldParentID;
				sbin.lpb = (LPBYTE) lpNotifications[i].info.obj.lpOldParentID;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.obj.lpOldParentID = \n"),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);
				sbin.cb = lpNotifications[i].info.obj.cbEntryID;
				sbin.lpb = (LPBYTE) lpNotifications[i].info.obj.lpEntryID;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.obj.lpEntryID = \n"),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);
				sbin.cb = lpNotifications[i].info.obj.cbParentID;
				sbin.lpb = (LPBYTE) lpNotifications[i].info.obj.lpParentID;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.obj.lpParentID = \n"),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.obj.ulObjType = 0x%08X"),i,
					lpNotifications[i].info.obj.ulObjType);

				InterpretFlags(PROP_ID(PR_OBJECT_TYPE),lpNotifications[i].info.obj.ulObjType, &szFlags);
				if (szFlags)
				{
					Outputf(ulDbgLvl,fFile,false,_T(" = %s"),
						szFlags);
				}
				delete[] szFlags;
				szFlags = NULL;

				Outputf(ulDbgLvl,fFile,false,_T("\n"));
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.obj.lpPropTagArray = \n"),i);
				_OutputPropTagArray(ulDbgLvl,fFile,
					lpNotifications[i].info.obj.lpPropTagArray);
			}
			break;
		case(fnevIndexing) :
			{
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.ext.ulEvent = 0x%08X\n"),i,
					lpNotifications[i].info.ext.ulEvent);
				SBinary sbin = {0};
				sbin.cb = lpNotifications[i].info.ext.cb;
				sbin.lpb = lpNotifications[i].info.ext.pbEventParameters;
				Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.ext.pbEventParameters = \n"),i);
				_OutputBinary(ulDbgLvl,fFile,&sbin);
				if (INDEXING_SEARCH_OWNER == lpNotifications[i].info.ext.ulEvent &&
					sizeof(INDEX_SEARCH_PUSHER_PROCESS) == lpNotifications[i].info.ext.cb)
				{
					Outputf(ulDbgLvl,fFile,true,_T("lpNotifications[%d].info.ext.ulEvent = INDEXING_SEARCH_OWNER\n"),i);

					INDEX_SEARCH_PUSHER_PROCESS* lpidxExt = (INDEX_SEARCH_PUSHER_PROCESS*)lpNotifications[i].info.ext.pbEventParameters;
					if (lpidxExt)
					{
						Outputf(ulDbgLvl,fFile,true,_T("lpidxExt->dwPID = 0x%08X\n"),lpidxExt->dwPID);
					}
				}
			}
			break;
		}
	}
	Outputf(ulDbgLvl,fFile,true,_T("End dumping notifications.\n"));
} // _OutputNotifications

void _OutputProperty(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropValue lpProp, _In_opt_ LPMAPIPROP lpObj)
{
	CHKPARAM;
	EARLYABORT;

	CString PropString;
	CString AltPropString;
	CString PropType;

	if (!lpProp)
	{
		_Output(ulDbgLvl,fFile, true,_T("OutputProperty called with NULL lpProp!\n"));
		return;
	}

	_Output(ulDbgLvl,fFile, true, TagToString(lpProp->ulPropTag,lpObj,false,false));

	_Output(ulDbgLvl,fFile, false, _T(", \""));
	if (PROP_TYPE(lpProp->ulPropTag) == PT_BINARY)
	{
		PropString = BinToHexString(&lpProp->Value.bin,true);
		AltPropString = BinToTextString(&lpProp->Value.bin,true);
	}
	else
	{
		InterpretProp(lpProp,&PropString,&AltPropString);
	}
	_Output(ulDbgLvl,fFile, false, (LPCTSTR) PropString);
	_Output(ulDbgLvl,fFile, false, _T("\", \""));
	_Output(ulDbgLvl,fFile, false, (LPCTSTR) AltPropString);

	LPTSTR szFlags = NULL;
	InterpretFlags(lpProp, &szFlags);
	if (szFlags)
	{
		_Output(ulDbgLvl,fFile, false, _T("\", \""));
		_Output(ulDbgLvl,fFile, false, szFlags);
	}
	delete[] szFlags;
	szFlags = NULL;

	_Output(ulDbgLvl,fFile, false, _T("\"\n"));
} // _OutputProperty

void _OutputProperties(ULONG ulDbgLvl, _In_opt_ FILE* fFile, ULONG cProps, _In_count_(cProps) LPSPropValue lpProps, _In_opt_ LPMAPIPROP lpObj)
{
	CHKPARAM;
	EARLYABORT;
	ULONG i;

	if (cProps && !lpProps)
	{
		_Output(ulDbgLvl, fFile, true, _T("OutputProperties called with NULL lpProps!\n"));
		return;
	}

	for (i = 0;i<cProps;i++)
	{
		_OutputProperty(ulDbgLvl, fFile, &lpProps[i], lpObj);
	}
} // _OutputProperties

void _OutputSRow(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRow lpSRow, _In_opt_ LPMAPIPROP lpObj)
{
	CHKPARAM;
	EARLYABORT;

	if (!lpSRow)
	{
		_Output(ulDbgLvl, fFile, true, _T("OutputSRow called with NULL lpSRow!\n"));
		return;
	}

	if (lpSRow->cValues && !lpSRow->lpProps)
	{
		_Output(ulDbgLvl, fFile, true, _T("OutputSRow called with NULL lpSRow->lpProps!\n"));
		return;
	}

	_OutputProperties(ulDbgLvl, fFile, lpSRow->cValues, lpSRow->lpProps,lpObj);
} // _OutputSRow

void _OutputSRowSet(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRowSet lpRowSet, _In_opt_ LPMAPIPROP lpObj)
{
	CHKPARAM;
	EARLYABORT;

	if (!lpRowSet)
	{
		_Output(ulDbgLvl, fFile, true, _T("OutputSRowSet called with NULL lpRowSet!\n"));
		return;
	}

	if (lpRowSet->cRows >= 1)
	{
		for (ULONG i = 0 ; i < lpRowSet->cRows ; i++)
		{
			_OutputSRow(ulDbgLvl,fFile,&lpRowSet->aRow[i],lpObj);
		}
	}
} // _OutputSRowSet

void _OutputRestriction(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj)
{
	CHKPARAM;
	EARLYABORT;

	if (!lpRes)
	{
		_Output(ulDbgLvl, fFile, true, _T("_OutputRestriction called with NULL lpRes!\n"));
		return;
	}

	_Output(ulDbgLvl, fFile, true, RestrictionToString(lpRes,lpObj));
} // _OutputRestriction

#define MAXBYTES 256
void _OutputStream(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSTREAM lpStream)
{
	CHKPARAM;
	EARLYABORT;
	HRESULT			hRes = S_OK;
	BYTE			bBuf[MAXBYTES+2]; // Allocate some extra for NULL terminators - 2 for Unicode
	ULONG			ulNumBytes = 0;
	LARGE_INTEGER	li = {0};

	if (!lpStream)
	{
		_Output(ulDbgLvl, fFile, true, _T("OutputStream called with NULL lpStream!\n"));
		return;
	}

	WC_H_MSG(lpStream->Seek(
		li,
		STREAM_SEEK_SET,
		NULL),IDS_STREAMSEEKFAILED);

	if (S_OK == hRes) do
	{
		hRes = S_OK;
		ulNumBytes = 0;
		EC_H(lpStream->Read(
			bBuf,
			MAXBYTES,
			&ulNumBytes));

		if (ulNumBytes >0)
		{
			bBuf[ulNumBytes] = 0;
			bBuf[ulNumBytes+1] = 0; // In case we are in Unicode
			_Output(ulDbgLvl, fFile, true, (TCHAR*)bBuf);
		}
	}
	while (ulNumBytes > 0);

} // _OutputStream

void _OutputVersion(ULONG ulDbgLvl, _In_opt_ FILE* fFile)
{
	CHKPARAM;
	EARLYABORT;
	TCHAR szFullPath[256];
	HRESULT hRes = S_OK;
	DWORD dwRet = 0;

	// Get version information from the application.
	EC_D(dwRet,GetModuleFileName(NULL, szFullPath, _countof(szFullPath)));

	if (S_OK == hRes)
	{
		DWORD dwVerHnd = 0;
		DWORD dwVerInfoSize = 0;

		EC_D(dwVerInfoSize,GetFileVersionInfoSize(szFullPath, &dwVerHnd));

		if (dwVerInfoSize)
		{
			// If we were able to get the information, process it.
			BYTE* pbData = new BYTE[dwVerInfoSize];
			if (pbData == NULL) return;

			BOOL bRet = false;
			EC_D(bRet,GetFileVersionInfo(
				szFullPath,
				dwVerHnd,
				dwVerInfoSize,
				(void*)pbData));

			if (pbData)
			{
				struct LANGANDCODEPAGE {
					WORD wLanguage;
					WORD wCodePage;
				} *lpTranslate = {0};

				UINT	cbTranslate = 0;
				UINT	iCodePages = 0;
				TCHAR	szSubBlock[256];

				// Read the list of languages and code pages.
				EC_B(VerQueryValue(
					pbData,
					_T("\\VarFileInfo\\Translation"), // STRING_OK
					(LPVOID*)&lpTranslate,
					&cbTranslate));

				// Read the file description for each language and code page.

				if (S_OK == hRes && lpTranslate)
				{
					for (iCodePages=0; iCodePages < (cbTranslate/sizeof(struct LANGANDCODEPAGE)); iCodePages++)
					{
						hRes = S_OK;
						EC_H(StringCchPrintf(
							szSubBlock,
							_countof(szSubBlock),
							_T("\\StringFileInfo\\%04x%04x\\"), // STRING_OK
							lpTranslate[iCodePages].wLanguage,
							lpTranslate[iCodePages].wCodePage));

						size_t cchRoot = NULL;
						EC_H(StringCchLength(szSubBlock,256,&cchRoot));

						// Load all our strings
						for (int iVerString = IDS_VER_FIRST; iVerString <= IDS_VER_LAST; iVerString++)
						{
							UINT	cchVer = 0;
							TCHAR*	lpszVer = NULL;
							TCHAR	szVerString[256];
							hRes = S_OK;

							int iRet = 0;
							EC_D(iRet,LoadString(
								GetModuleHandle(NULL),
								iVerString,
								szVerString,
								_countof(szVerString)));

							EC_H(StringCchCopy(
								&szSubBlock[cchRoot],
								_countof(szSubBlock) - cchRoot,
								szVerString));

							EC_B(VerQueryValue(
								(void*)pbData,
								szSubBlock,
								(void**)&lpszVer,
								&cchVer));

							if (S_OK == hRes && cchVer && lpszVer)
							{
								Outputf(ulDbgLvl, fFile, true, _T("%s: %s\n"),szVerString,lpszVer);
							}
						}
					}
				}
			}
			delete[] pbData;
		}
	}
} // _OutputVersion

void OutputXMLValueToFile(_In_opt_ FILE* fFile, UINT uidTag, _In_opt_z_ LPCTSTR szValue, int iIndent)
{
	if (! fFile || !szValue || !uidTag) return;

	if (!szValue[0]) return;

	CString szTag;
	HRESULT hRes = S_OK;
	EC_B(szTag.LoadString(uidTag));

	int i;
	for (i = 0;i<iIndent;i++) OutputToFile(fFile,_T("\t"));
	OutputToFilef(fFile,_T("<%s>"),(LPCTSTR) szTag);
	OutputToFile(fFile,szValue);
	OutputToFilef(fFile,_T("</%s>\n"),(LPCTSTR) szTag);
} // OutputXMLValueToFile

void OutputCDataOpen(_In_opt_ FILE* fFile)
{
	OutputToFile(fFile,_T("<![CDATA["));
} // OutputCDataOpen

void OutputCDataClose(_In_opt_ FILE* fFile)
{
	OutputToFile(fFile,_T("]]>\n"));
} // OutputCDataClose

void OutputXMLCDataValueToFile(_In_opt_ FILE* fFile, UINT uidTag, _In_z_ LPCTSTR szValue, int iIndent)
{
	if (! fFile || !szValue || !uidTag) return;

	if (!szValue[0]) return;

	CString szTag;
	HRESULT hRes = S_OK;
	EC_B(szTag.LoadString(uidTag));

	int i;
	for (i = 0;i<iIndent;i++) OutputToFile(fFile,_T("\t"));
	OutputToFilef(fFile,_T("<%s>"),(LPCTSTR) szTag);
	OutputCDataOpen(fFile);
	OutputToFile(fFile,szValue);
	OutputCDataClose(fFile);
	for (i = 0;i<iIndent;i++) OutputToFile(fFile,_T("\t"));
	OutputToFilef(fFile,_T("</%s>\n"),(LPCTSTR) szTag);
} // OutputXMLCDataValueToFile