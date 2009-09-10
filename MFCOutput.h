#pragma once
// MFCOutput.h : header file
//
// Output (to File/Debug) functions

void OpenDebugFile();
void CloseDebugFile();
void SetDebugLevel(ULONG ulDbgLvl);
void SetDebugOutputToFile(BOOL bDoOutput);

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
#define DBGAddIn 							((ULONG) 0x00010000)
#define DBGStream							((ULONG) 0x00020000)
#define DBGMenu								((ULONG) 0x80000000)

// Super verbose is really overkill - scale back for our ALL default
#define DBGAll								((ULONG) 0x0000ffff)
#define DBGSuperVerbose						((ULONG) 0xffffffff)

#define fIsSet(ulTag)	(((ulTag) != DBGNoDebug) && (RegKeys[regkeyDEBUG_TAG].ulCurDWORD & (ulTag)))

FILE* OpenFile(LPCTSTR szFileName,BOOL bNewFile);
void CloseFile(FILE* fFile);

void _Output(ULONG ulDbgLvl, FILE* fFile, BOOL bPrintThreadTime, LPCTSTR szMsg);
void __cdecl Outputf(ULONG ulDbgLvl, FILE* fFile, BOOL bPrintThreadTime, LPCTSTR szMsg,...);

#define OutputToFile(fFile, szMsg) _Output((DBGNoDebug), (fFile), true, (szMsg))
void __cdecl OutputToFilef(FILE* fFile, LPCTSTR szMsg,...);

void __cdecl DebugPrint(ULONG ulDbgLvl,LPCTSTR szMsg,...);
void __cdecl DebugPrintEx(ULONG ulDbgLvl,LPCTSTR szClass, LPCTSTR szFunc, LPCTSTR szMsg, ...);

// Template for the Output functions
// void Output(ULONG ulDbgLvl, LPCTSTR szFileName,stufftooutput)

// If the first parameter is not DBGNoDebug, we debug print the output
// If the second parameter is a file name, we print the output to the file
// Doing both is OK, if we ask to do neither, ASSERT

// We'll use macros to make these calls so the code will read right

void _OutputBinary(ULONG ulDbgLvl, FILE* fFile, LPSBinary lpBin);
void _OutputProperty(ULONG ulDbgLvl, FILE* fFile, LPSPropValue lpProp, LPMAPIPROP lpObj);
void _OutputProperties(ULONG ulDbgLvl, FILE* fFile, ULONG cProps, LPSPropValue lpProps, LPMAPIPROP lpObj);
void _OutputSRow(ULONG ulDbgLvl, FILE* fFile, LPSRow lpSRow, LPMAPIPROP lpObj);
void _OutputSRowSet(ULONG ulDbgLvl, FILE* fFile, LPSRowSet lpRowSet, LPMAPIPROP lpObj);
void _OutputRestriction(ULONG ulDbgLvl, FILE* fFile, LPSRestriction lpRes, LPMAPIPROP lpObj);
void _OutputStream(ULONG ulDbgLvl, FILE* fFile, LPSTREAM lpStream);
void _OutputVersion(ULONG ulDbgLvl, FILE* fFile);
void _OutputFormInfo(ULONG ulDbgLvl, FILE* fFile, LPMAPIFORMINFO lpMAPIFormInfo);
void _OutputFormPropArray(ULONG ulDbgLvl, FILE* fFile, LPMAPIFORMPROPARRAY lpMAPIFormPropArray);
void _OutputPropTagArray(ULONG ulDbgLvl, FILE* fFile, LPSPropTagArray lpTagsToDump);
void _OutputTable(ULONG ulDbgLvl, FILE* fFile, LPMAPITABLE lpMAPITable);
void _OutputNotifications(ULONG ulDbgLvl, FILE* fFile, ULONG cNotify, LPNOTIFICATION lpNotifications);

#define DebugPrintBinary(ulDbgLvl, lpBin)						_OutputBinary((ulDbgLvl), NULL, (lpBin))
#define DebugPrintProperties(ulDbgLvl, cProps, lpProps, lpObj)	_OutputProperties((ulDbgLvl), NULL, (cProps), (lpProps), (lpObj))
#define DebugPrintRestriction(ulDbgLvl, lpRes, lpObj)			_OutputRestriction((ulDbgLvl), NULL, (lpRes), (lpObj))
#define DebugPrintStream(ulDbgLvl, lpStream)					_OutputStream((ulDbgLvl), NULL, lpStream)
#define DebugPrintVersion(ulDbgLvl)								_OutputVersion((ulDbgLvl), NULL)
#define DebugPrintFormInfo(ulDbgLvl,lpMAPIFormInfo)				_OutputFormInfo((ulDbgLvl),NULL, (lpMAPIFormInfo))
#define DebugPrintFormPropArray(ulDbgLvl,lpMAPIFormPropArray)	_OutputFormPropArray((ulDbgLvl),NULL, (lpMAPIFormPropArray))
#define DebugPrintPropTagArray(ulDbgLvl,lpTagsToDump)			_OutputPropTagArray((ulDbgLvl),NULL, (lpTagsToDump))
#define DebugPrintNotifications(ulDbgLvl, cNotify, lpNotifications)		_OutputNotifications((ulDbgLvl),NULL, (cNotify), (lpNotifications))
#define DebugPrintSRowSet(ulDbgLvl, lpRowSet, lpObj)			_OutputSRowSet((ulDbgLvl), NULL, (lpRowSet), (lpObj))

#define OutputStreamToFile(fFile, lpStream)					_OutputStream(DBGNoDebug, (fFile), (lpStream))
#define OutputTableToFile(fFile, lpMAPITable)					_OutputTable(DBGNoDebug, (fFile), (lpMAPITable))

// We'll only output this information in debug builds.
#ifdef _DEBUG
#define TRACE_CONSTRUCTOR(__class) DebugPrintEx(DBGConDes,(__class),(__class),_T("(this = 0x%X) - Constructor\n"),this);
#define TRACE_DESTRUCTOR(__class) DebugPrintEx(DBGConDes,(__class),(__class),_T("(this = 0x%X) - Destructor\n"),this);

#define TRACE_ADDREF(__class,__count) DebugPrintEx(DBGRefCount,(__class),_T("AddRef"),_T("(this = 0x%X) m_cRef increased to %d.\n"),this,(__count));
#define TRACE_RELEASE(__class,__count) DebugPrintEx(DBGRefCount,(__class),_T("Release"),_T("(this = 0x%X) m_cRef decreased to %d.\n"),this,(__count));
#else
#define TRACE_CONSTRUCTOR(__class)
#define TRACE_DESTRUCTOR(__class)

#define TRACE_ADDREF(__class,__count)
#define TRACE_RELEASE(__class,__count)
#endif

void OutputXMLValueToFile(FILE* fFile,UINT uidTag, LPCTSTR szValue, int iIndent);
void OutputCDataOpen(FILE* fFile);
void OutputCDataClose(FILE* fFile);
void OutputXMLCDataValueToFile(FILE* fFile,UINT uidTag, LPCTSTR szValue, int iIndent);
