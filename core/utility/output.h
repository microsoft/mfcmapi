#pragma once
// Output (to File/Debug) functions

namespace output
{
	extern std::wstring g_szXMLHeader;
	extern std::function<void(const std::wstring& errString)> outputToDbgView;

	void OpenDebugFile();
	void CloseDebugFile();
	_Check_return_ ULONG GetDebugLevel();
	void SetDebugLevel(ULONG ulDbgLvl);
	void SetDebugOutputToFile(bool bDoOutput);

	// New system for debug output: When outputting debug output, a tag is included - if that tag is
	// set in registry::debugTag, then we do the output. Otherwise, we ditch it.
	// DBGNoDebug never gets output - special case

	// The global debug level - combination of flags from below
	// registry::debugTag

#define DBGNoDebug ((ULONG) 0x00000000)
#define DBGGeneric ((ULONG) 0x00000001)
#define DBGVersionBanner ((ULONG) 0x00000002)
#define DBGFatalError ((ULONG) 0x00000004)
#define DBGRefCount ((ULONG) 0x00000008)
#define DBGConDes ((ULONG) 0x00000010)
#define DBGNotify ((ULONG) 0x00000020)
#define DBGHRes ((ULONG) 0x00000040)
#define DBGCreateDialog ((ULONG) 0x00000080)
#define DBGOpenItemProp ((ULONG) 0x00000100)
#define DBGDeleteSelectedItem ((ULONG) 0x00000200)
#define DBGTest ((ULONG) 0x00000400)
#define DBGFormViewer ((ULONG) 0x00000800)
#define DBGNamedProp ((ULONG) 0x00001000)
#define DBGLoadLibrary ((ULONG) 0x00002000)
#define DBGForms ((ULONG) 0x00004000)
#define DBGAddInPlumbing ((ULONG) 0x00008000)
#define DBGAddIn ((ULONG) 0x00010000)
#define DBGStream ((ULONG) 0x00020000)
#define DBGSmartView ((ULONG) 0x00040000)
#define DBGLoadMAPI ((ULONG) 0x00080000)
#define DBGHierarchy ((ULONG) 0x00100000)
#define DBGNamedPropCacheMisses ((ULONG) 0x00200000)
#define DBGDraw ((ULONG) 0x10000000)
#define DBGUI ((ULONG) 0x20000000)
#define DBGMAPIFunctions ((ULONG) 0x40000000)
#define DBGMenu ((ULONG) 0x80000000)

// Super verbose is really overkill - scale back for our ALL default
#define DBGAll ((ULONG) 0x0000ffff)
#define DBGSuperVerbose ((ULONG) 0xffffffff)

#define fIsSet(ulTag) (registry::debugTag & (ulTag))
#define fIsSetv(ulTag) (((ulTag) != DBGNoDebug) && (registry::debugTag & (ulTag)))

#define CHKPARAM assert(DBGNoDebug != ulDbgLvl || fFile)

	// quick check to see if we have anything to print - so we can avoid executing the call
#define EARLYABORT \
	{ \
		if (!fFile && !registry::debugToFile && !fIsSetv(ulDbgLvl)) return; \
	}

	_Check_return_ FILE* MyOpenFile(const std::wstring& szFileName, bool bNewFile);
	_Check_return_ FILE* MyOpenFileMode(const std::wstring& szFileName, const wchar_t* mode);
	void CloseFile(_In_opt_ FILE* fFile);

	void Output(ULONG ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, const std::wstring& szMsg);
	void __cdecl Outputf(ULONG ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef Outputf
#define Outputf(ulDbgLvl, fFile, bPrintThreadTime, szMsg, ...) \
	Outputf((wprintf(szMsg, __VA_ARGS__), ulDbgLvl), fFile, bPrintThreadTime, szMsg, __VA_ARGS__)
#endif

#define OutputToFile(fFile, szMsg) Output((DBGNoDebug), (fFile), true, (szMsg))
	void __cdecl OutputToFilef(_In_opt_ FILE* fFile, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef OutputToFilef
#define OutputToFilef(fFile, szMsg, ...) OutputToFilef((wprintf(szMsg, __VA_ARGS__), fFile), szMsg, __VA_ARGS__)
#endif

	void __cdecl DebugPrint(ULONG ulDbgLvl, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef DebugPrint
#define DebugPrint(ulDbgLvl, szMsg, ...) DebugPrint((wprintf(szMsg, __VA_ARGS__), ulDbgLvl), szMsg, __VA_ARGS__)
#endif

	void __cdecl DebugPrintEx(ULONG ulDbgLvl, std::wstring& szClass, const std::wstring& szFunc, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef DebugPrintEx
#define DebugPrintEx(ulDbgLvl, szClass, szFunc, szMsg, ...) \
	DebugPrintEx((wprintf(szMsg, __VA_ARGS__), ulDbgLvl), szClass, szFunc, szMsg, __VA_ARGS__)
#endif

	// We'll only output this information in debug builds.
#ifdef _DEBUG
#define TRACE_CONSTRUCTOR(__class) \
	output::DebugPrintEx(DBGConDes, (__class), (__class), L"(this = %p) - Constructor\n", this);
#define TRACE_DESTRUCTOR(__class) \
	output::DebugPrintEx(DBGConDes, (__class), (__class), L"(this = %p) - Destructor\n", this);

#define TRACE_ADDREF(__class, __count) \
	output::DebugPrintEx(DBGRefCount, (__class), L"AddRef", L"(this = %p) m_cRef increased to %d.\n", this, (__count));
#define TRACE_RELEASE(__class, __count) \
	output::DebugPrintEx(DBGRefCount, (__class), L"Release", L"(this = %p) m_cRef decreased to %d.\n", this, (__count));
#else
#define TRACE_CONSTRUCTOR(__class)
#define TRACE_DESTRUCTOR(__class)

#define TRACE_ADDREF(__class, __count)
#define TRACE_RELEASE(__class, __count)
#endif

	void _OutputVersion(ULONG ulDbgLvl, _In_opt_ FILE* fFile);
	void _OutputStream(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSTREAM lpStream);
#define DebugPrintVersion(ulDbgLvl) _OutputVersion((ulDbgLvl), nullptr)
#define DebugPrintStream(ulDbgLvl, lpStream) _OutputStream((ulDbgLvl), nullptr, lpStream)

	void OutputXMLValue(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		UINT uidTag,
		const std::wstring& szValue,
		bool bWrapCData,
		int iIndent);
	void OutputCDataOpen(ULONG ulDbgLvl, _In_opt_ FILE* fFile);
	void OutputCDataClose(ULONG ulDbgLvl, _In_opt_ FILE* fFile);

#define OutputXMLValueToFile(fFile, uidTag, szValue, bWrapCData, iIndent) \
	OutputXMLValue(DBGNoDebug, fFile, uidTag, szValue, bWrapCData, iIndent)
} // namespace output