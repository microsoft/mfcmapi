#pragma once
// MFCOutput.h : header file
//
// Output (to File/Debug) functions

extern LPCTSTR g_szXMLHeader;

void OpenDebugFile();
void CloseDebugFile();
_Check_return_ ULONG GetDebugLevel();
void SetDebugLevel(ULONG ulDbgLvl);
void SetDebugOutputToFile(bool bDoOutput);

// New system for debug output: When outputting debug output, a tag is included - if that tag is
// set in RegKeys[regkeyDEBUG_TAG].ulCurDWORD, then we do the output. Otherwise, we ditch it.
// DBGNoDebug never gets output - special case

// The global debug level - combination of flags from below
// RegKeys[regkeyDEBUG_TAG].ulCurDWORD

#define	DBGNoDebug							((ULONG) 0x00000000)
#define	DBGGeneric							((ULONG) 0x00000001)
#define	DBGVersionBanner					((ULONG) 0x00000002)
#define	DBGFatalError						((ULONG) 0x00000004)
#define	DBGRefCount							((ULONG) 0x00000008)
#define	DBGConDes							((ULONG) 0x00000010)
#define	DBGNotify							((ULONG) 0x00000020)
#define	DBGHRes								((ULONG) 0x00000040)
#define DBGCreateDialog						((ULONG) 0x00000080)
#define DBGOpenItemProp						((ULONG) 0x00000100)
#define DBGDeleteSelectedItem				((ULONG) 0x00000200)
#define DBGTest								((ULONG) 0x00000400)
#define DBGFormViewer						((ULONG) 0x00000800)
#define DBGNamedProp						((ULONG) 0x00001000)
#define DBGLoadLibrary						((ULONG) 0x00002000)
#define DBGForms							((ULONG) 0x00004000)
#define DBGAddInPlumbing					((ULONG) 0x00008000)
#define DBGAddIn							((ULONG) 0x00010000)
#define DBGStream							((ULONG) 0x00020000)
#define DBGSmartView						((ULONG) 0x00040000)
#define DBGLoadMAPI							((ULONG) 0x00080000)
#define DBGHierarchy						((ULONG) 0x00100000)
#define DBGMAPIFunctions					((ULONG) 0x40000000)
#define DBGMenu								((ULONG) 0x80000000)

// Super verbose is really overkill - scale back for our ALL default
#define DBGAll								((ULONG) 0x0000ffff)
#define DBGSuperVerbose						((ULONG) 0xffffffff)

#define fIsSet(ulTag) (RegKeys[regkeyDEBUG_TAG].ulCurDWORD & (ulTag))
#define fIsSetv(ulTag) (((ulTag) != DBGNoDebug) && (RegKeys[regkeyDEBUG_TAG].ulCurDWORD & (ulTag)))

_Check_return_ FILE* MyOpenFile(_In_z_ LPCWSTR szFileName, bool bNewFile);
void CloseFile(_In_opt_ FILE* fFile);

void _Output(ULONG ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, _In_opt_z_ LPCTSTR szMsg);
void __cdecl Outputf(ULONG ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, _Printf_format_string_ LPCTSTR szMsg, ...);

#define OutputToFile(fFile, szMsg) _Output((DBGNoDebug), (fFile), true, (szMsg))
void __cdecl OutputToFilef(_In_opt_ FILE* fFile, _Printf_format_string_ LPCTSTR szMsg, ...);

void __cdecl DebugPrint(ULONG ulDbgLvl, _Printf_format_string_ LPCTSTR szMsg, ...);
void __cdecl DebugPrintEx(ULONG ulDbgLvl, _In_z_ LPCTSTR szClass, _In_z_ LPCTSTR szFunc, _Printf_format_string_ LPCTSTR szMsg, ...);

// Template for the Output functions
// void Output(ULONG ulDbgLvl, LPCTSTR szFileName,stufftooutput)

// If the first parameter is not DBGNoDebug, we debug print the output
// If the second parameter is a file name, we print the output to the file
// Doing both is OK, if we ask to do neither, ASSERT

// We'll use macros to make these calls so the code will read right

void _OutputBinary(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSBinary lpBin);
void _OutputProperty(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropValue lpProp, _In_opt_ LPMAPIPROP lpObj, bool bRetryStreamProps);
void _OutputProperties(ULONG ulDbgLvl, _In_opt_ FILE* fFile, ULONG cProps, _In_count_(cProps) LPSPropValue lpProps, _In_opt_ LPMAPIPROP lpObj, bool bRetryStreamProps);
void _OutputSRow(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRow lpSRow, _In_opt_ LPMAPIPROP lpObj);
void _OutputSRowSet(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRowSet lpRowSet, _In_opt_ LPMAPIPROP lpObj);
void _OutputRestriction(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_opt_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj);
void _OutputStream(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSTREAM lpStream);
void _OutputVersion(ULONG ulDbgLvl, _In_opt_ FILE* fFile);
void _OutputFormInfo(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMINFO lpMAPIFormInfo);
void _OutputFormPropArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMPROPARRAY lpMAPIFormPropArray);
void _OutputPropTagArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropTagArray lpTagsToDump);
void _OutputTable(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPITABLE lpMAPITable);
void _OutputNotifications(ULONG ulDbgLvl, _In_opt_ FILE* fFile, ULONG cNotify, _In_count_(cNotify) LPNOTIFICATION lpNotifications, _In_opt_ LPMAPIPROP lpObj);
void _OutputEntryList(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPENTRYLIST lpEntryList);

#define DebugPrintBinary(ulDbgLvl, lpBin)						_OutputBinary((ulDbgLvl), NULL, (lpBin))
#define DebugPrintProperties(ulDbgLvl, cProps, lpProps, lpObj)	_OutputProperties((ulDbgLvl), NULL, (cProps), (lpProps), (lpObj), false)
#define DebugPrintRestriction(ulDbgLvl, lpRes, lpObj)			_OutputRestriction((ulDbgLvl), NULL, (lpRes), (lpObj))
#define DebugPrintStream(ulDbgLvl, lpStream)					_OutputStream((ulDbgLvl), NULL, lpStream)
#define DebugPrintVersion(ulDbgLvl)								_OutputVersion((ulDbgLvl), NULL)
#define DebugPrintFormInfo(ulDbgLvl,lpMAPIFormInfo)				_OutputFormInfo((ulDbgLvl),NULL, (lpMAPIFormInfo))
#define DebugPrintFormPropArray(ulDbgLvl,lpMAPIFormPropArray)	_OutputFormPropArray((ulDbgLvl),NULL, (lpMAPIFormPropArray))
#define DebugPrintPropTagArray(ulDbgLvl,lpTagsToDump)			_OutputPropTagArray((ulDbgLvl),NULL, (lpTagsToDump))
#define DebugPrintNotifications(ulDbgLvl, cNotify, lpNotifications, lpObj)		_OutputNotifications((ulDbgLvl),NULL, (cNotify), (lpNotifications), (lpObj))
#define DebugPrintSRowSet(ulDbgLvl, lpRowSet, lpObj)			_OutputSRowSet((ulDbgLvl), NULL, (lpRowSet), (lpObj))
#define DebugPrintEntryList(ulDbgLvl, lpEntryList)				_OutputEntryList((ulDbgLvl), NULL, (lpEntryList))

#define OutputStreamToFile(fFile, lpStream)						_OutputStream(DBGNoDebug, (fFile), (lpStream))
#define OutputTableToFile(fFile, lpMAPITable)					_OutputTable(DBGNoDebug, (fFile), (lpMAPITable))
#define OutputSRowToFile(fFile, lpSRow, lpObj)					_OutputSRow(DBGNoDebug,fFile, lpSRow, lpObj)
#define OutputPropertiesToFile(fFile,cProps,lpProps,lpObj,bRetry) _OutputProperties(DBGNoDebug,fFile, cProps, lpProps, lpObj, bRetry)
#define OutputPropertyToFile(fFile, lpProp, lpObj, bRetry)		_OutputProperty(DBGNoDebug,fFile, lpProp, lpObj, bRetry)

// We'll only output this information in debug builds.
#ifdef _DEBUG
#define TRACE_CONSTRUCTOR(__class) DebugPrintEx(DBGConDes,(__class),(__class),_T("(this = %p) - Constructor\n"),this);
#define TRACE_DESTRUCTOR(__class) DebugPrintEx(DBGConDes,(__class),(__class),_T("(this = %p) - Destructor\n"),this);

#define TRACE_ADDREF(__class,__count) DebugPrintEx(DBGRefCount,(__class),_T("AddRef"),_T("(this = %p) m_cRef increased to %d.\n"),this,(__count));
#define TRACE_RELEASE(__class,__count) DebugPrintEx(DBGRefCount,(__class),_T("Release"),_T("(this = %p) m_cRef decreased to %d.\n"),this,(__count));
#else
#define TRACE_CONSTRUCTOR(__class)
#define TRACE_DESTRUCTOR(__class)

#define TRACE_ADDREF(__class,__count)
#define TRACE_RELEASE(__class,__count)
#endif

void OutputXMLValue(ULONG ulDbgLvl, _In_opt_ FILE* fFile, UINT uidTag, _In_z_ LPTSTR szValue, bool bWrapCData, int iIndent);
void OutputCDataOpen(ULONG ulDbgLvl, _In_opt_ FILE* fFile);
void OutputCDataClose(ULONG ulDbgLvl, _In_opt_ FILE* fFile);

#define OutputXMLValueToFile(fFile, uidTag, szValue, bWrapCData, iIndent) OutputXMLValue(DBGNoDebug, fFile, uidTag, szValue, bWrapCData, iIndent)