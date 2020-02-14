#pragma once
// Output (to File/Debug) functions

namespace output
{
	// New system for debug output: When outputting debug output, a tag is included - if that tag is
	// set in registry::debugTag, then we do the output. Otherwise, we ditch it.
	// DBGNoDebug never gets output - special case

	// The global debug level - combination of flags from below
	// registry::debugTag
	enum class dbgLevel : UINT
	{
		NoDebug = 0x00000000,
		Generic = 0x00000001,
		VersionBanner = 0x00000002,
		FatalError = 0x00000004,
		RefCount = 0x00000008,
		ConDes = 0x00000010,
		Notify = 0x00000020,
		HRes = 0x00000040,
		CreateDialog = 0x00000080,
		OpenItemProp = 0x00000100,
		DeleteSelectedItem = 0x00000200,
		Test = 0x00000400,
		FormViewer = 0x00000800,
		NamedProp = 0x00001000,
		LoadLibrary = 0x00002000,
		Forms = 0x00004000,
		AddInPlumbing = 0x00008000,
		AddIn = 0x00010000,
		Stream = 0x00020000,
		SmartView = 0x00040000,
		LoadMAPI = 0x00080000,
		Hierarchy = 0x00100000,
		NamedPropCacheMisses = 0x00200000,
		Draw = 0x10000000,
		UI = 0x20000000,
		MAPIFunctions = 0x40000000,
		Menu = 0x80000000,

		// Super verbose is really overkill - scale back for our ALL default
		All = 0x0000ffff,
		SuperVerbose = 0xffffffff,
	};

	extern std::wstring g_szXMLHeader;
	extern std::function<void(const std::wstring& errString)> outputToDbgView;

	void OpenDebugFile();
	void CloseDebugFile();
	void SetDebugOutputToFile(bool bDoOutput);

	bool fIsSet(dbgLevel ulTag) noexcept;
	bool fIsSetv(dbgLevel ulTag) noexcept;
	bool earlyExit(dbgLevel ulDbgLvl, bool fFile);

	_Check_return_ FILE* MyOpenFile(const std::wstring& szFileName, bool bNewFile);
	_Check_return_ FILE* MyOpenFileMode(const std::wstring& szFileName, const wchar_t* mode);
	void CloseFile(_In_opt_ FILE* fFile) noexcept;

	void Output(dbgLevel ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, const std::wstring& szMsg);
	void __cdecl Outputf(dbgLevel ulDbgLvl, _In_opt_ FILE* fFile, bool bPrintThreadTime, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef Outputf
#define Outputf(ulDbgLvl, fFile, bPrintThreadTime, szMsg, ...) \
	Outputf((wprintf(szMsg, __VA_ARGS__), ulDbgLvl), fFile, bPrintThreadTime, szMsg, __VA_ARGS__)
#endif

#define OutputToFile(fFile, szMsg) Output((output::dbgLevel::NoDebug), (fFile), true, (szMsg))
	void __cdecl OutputToFilef(_In_opt_ FILE* fFile, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef OutputToFilef
#define OutputToFilef(fFile, szMsg, ...) OutputToFilef((wprintf(szMsg, __VA_ARGS__), fFile), szMsg, __VA_ARGS__)
#endif

	void __cdecl DebugPrint(dbgLevel ulDbgLvl, LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef DebugPrint
#define DebugPrint(ulDbgLvl, szMsg, ...) DebugPrint((wprintf(szMsg, __VA_ARGS__), ulDbgLvl), szMsg, __VA_ARGS__)
#endif
	void __cdecl DebugPrint(dbgLevel ulDbgLvl, LPCWSTR szMsg, va_list argList);

	void __cdecl DebugPrintEx(
		dbgLevel ulDbgLvl,
		const std::wstring& szClass,
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
	output::DebugPrintEx(output::dbgLevel::ConDes, (__class), (__class), L"(this = %p) - Constructor\n", this);
#define TRACE_DESTRUCTOR(__class) \
	output::DebugPrintEx(output::dbgLevel::ConDes, (__class), (__class), L"(this = %p) - Destructor\n", this);

#define TRACE_ADDREF(__class, __count) \
	output::DebugPrintEx( \
		output::dbgLevel::RefCount, (__class), L"AddRef", L"(this = %p) m_cRef increased to %d.\n", this, (__count));
#define TRACE_RELEASE(__class, __count) \
	output::DebugPrintEx( \
		output::dbgLevel::RefCount, (__class), L"Release", L"(this = %p) m_cRef decreased to %d.\n", this, (__count));
#else
#define TRACE_CONSTRUCTOR(__class)
#define TRACE_DESTRUCTOR(__class)

#define TRACE_ADDREF(__class, __count)
#define TRACE_RELEASE(__class, __count)
#endif

	void outputVersion(dbgLevel ulDbgLvl, _In_opt_ FILE* fFile);
	void outputStream(dbgLevel ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSTREAM lpStream);

	void OutputXMLValue(
		dbgLevel ulDbgLvl,
		_In_opt_ FILE* fFile,
		UINT uidTag,
		const std::wstring& szValue,
		bool bWrapCData,
		int iIndent);
	void OutputCDataOpen(dbgLevel ulDbgLvl, _In_opt_ FILE* fFile);
	void OutputCDataClose(dbgLevel ulDbgLvl, _In_opt_ FILE* fFile);

#define OutputXMLValueToFile(fFile, uidTag, szValue, bWrapCData, iIndent) \
	OutputXMLValue(output::dbgLevel::NoDebug, fFile, uidTag, szValue, bWrapCData, iIndent)

	void initStubCallbacks();
} // namespace output