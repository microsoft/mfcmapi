#pragma once
// Output (to File/Debug) functions

namespace output
{
	// New system for debug output: When outputting debug output, a tag is included - if that tag is
	// set in registry::debugTag, then we do the output. Otherwise, we ditch it.
	// DBGNoDebug never gets output - special case

	// The global debug level - combination of flags from below
	// registry::debugTag
	enum DBGLEVEL
	{
		DBGNoDebug = 0x00000000,
		DBGGeneric = 0x00000001,
		DBGVersionBanner = 0x00000002,
		DBGFatalError = 0x00000004,
		DBGRefCount = 0x00000008,
		DBGConDes = 0x00000010,
		DBGNotify = 0x00000020,
		DBGHRes = 0x00000040,
		DBGCreateDialog = 0x00000080,
		DBGOpenItemProp = 0x00000100,
		DBGDeleteSelectedItem = 0x00000200,
		DBGTest = 0x00000400,
		DBGFormViewer = 0x00000800,
		DBGNamedProp = 0x00001000,
		DBGLoadLibrary = 0x00002000,
		DBGForms = 0x00004000,
		DBGAddInPlumbing = 0x00008000,
		DBGAddIn = 0x00010000,
		DBGStream = 0x00020000,
		DBGSmartView = 0x00040000,
		DBGLoadMAPI = 0x00080000,
		DBGHierarchy = 0x00100000,
		DBGNamedPropCacheMisses = 0x00200000,
		DBGDraw = 0x10000000,
		DBGUI = 0x20000000,
		DBGMAPIFunctions = 0x40000000,
		DBGMenu = 0x80000000,

		// Super verbose is really overkill - scale back for our ALL default
		DBGAll = 0x0000ffff,
		DBGSuperVerbose = 0xffffffff,
	};

	extern std::wstring g_szXMLHeader;
	extern std::function<void(const std::wstring& errString)> outputToDbgView;

	void OpenDebugFile();
	void CloseDebugFile();
	void SetDebugOutputToFile(bool bDoOutput);

	bool fIsSet(output::DBGLEVEL ulTag);
	bool fIsSetv(output::DBGLEVEL ulTag);
	bool earlyExit(output::DBGLEVEL ulDbgLvl, bool fFile);

	_Check_return_ FILE* MyOpenFile(const std::wstring& szFileName, bool bNewFile);
	_Check_return_ FILE* MyOpenFileMode(const std::wstring& szFileName, const wchar_t* mode);
	void CloseFile(_In_opt_ FILE* fFile);

	void Output(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, const std::wstring& szMsg);
	void __cdecl Outputf(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef Outputf
#define Outputf(ulDbgLvl, fFile, bPrintThreadTime, szMsg, ...) \
	Outputf((wprintf(szMsg, __VA_ARGS__), ulDbgLvl), fFile, bPrintThreadTime, szMsg, __VA_ARGS__)
#endif

#define OutputToFile(fFile, szMsg) Output((output::DBGNoDebug), (fFile), true, (szMsg))
	void __cdecl OutputToFilef(_In_opt_ FILE* fFile, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef OutputToFilef
#define OutputToFilef(fFile, szMsg, ...) OutputToFilef((wprintf(szMsg, __VA_ARGS__), fFile), szMsg, __VA_ARGS__)
#endif

	void __cdecl DebugPrint(output::DBGLEVEL ulDbgLvl, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef DebugPrint
#define DebugPrint(ulDbgLvl, szMsg, ...) DebugPrint((wprintf(szMsg, __VA_ARGS__), ulDbgLvl), szMsg, __VA_ARGS__)
#endif
	void __cdecl DebugPrint(output::DBGLEVEL ulDbgLvl, LPCWSTR szMsg, va_list argList);

	void __cdecl DebugPrintEx(
		output::DBGLEVEL ulDbgLvl,
		std::wstring& szClass,
		const std::wstring& szFunc,
		LPCWSTR szMsg,
		...);
#ifdef CHECKFORMATPARAMS
#undef DebugPrintEx
#define DebugPrintEx(ulDbgLvl, szClass, szFunc, szMsg, ...) \
	DebugPrintEx((wprintf(szMsg, __VA_ARGS__), ulDbgLvl), szClass, szFunc, szMsg, __VA_ARGS__)
#endif

	// We'll only output this information in debug builds.
#ifdef _DEBUG
#define TRACE_CONSTRUCTOR(__class) \
	output::DebugPrintEx(output::DBGConDes, (__class), (__class), L"(this = %p) - Constructor\n", this);
#define TRACE_DESTRUCTOR(__class) \
	output::DebugPrintEx(output::DBGConDes, (__class), (__class), L"(this = %p) - Destructor\n", this);

#define TRACE_ADDREF(__class, __count) \
	output::DebugPrintEx( \
		output::DBGRefCount, (__class), L"AddRef", L"(this = %p) m_cRef increased to %d.\n", this, (__count));
#define TRACE_RELEASE(__class, __count) \
	output::DebugPrintEx( \
		output::DBGRefCount, (__class), L"Release", L"(this = %p) m_cRef decreased to %d.\n", this, (__count));
#else
#define TRACE_CONSTRUCTOR(__class)
#define TRACE_DESTRUCTOR(__class)

#define TRACE_ADDREF(__class, __count)
#define TRACE_RELEASE(__class, __count)
#endif

	void outputVersion(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile);
	void outputStream(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSTREAM lpStream);

	void OutputXMLValue(
		output::DBGLEVEL ulDbgLvl,
		_In_opt_ FILE* fFile,
		UINT uidTag,
		const std::wstring& szValue,
		bool bWrapCData,
		int iIndent);
	void OutputCDataOpen(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile);
	void OutputCDataClose(output::DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile);

#define OutputXMLValueToFile(fFile, uidTag, szValue, bWrapCData, iIndent) \
	OutputXMLValue(output::DBGNoDebug, fFile, uidTag, szValue, bWrapCData, iIndent)
} // namespace output