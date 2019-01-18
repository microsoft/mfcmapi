#pragma once

// Forward declarations
enum __ParsingTypeEnum;
struct NAME_ARRAY_ENTRY_V2;
struct NAME_ARRAY_ENTRY;
struct GUID_ARRAY_ENTRY;
struct NAMEID_ARRAY_ENTRY;
struct FLAG_ARRAY_ENTRY;
struct SMARTVIEW_PARSER_ARRAY_ENTRY;
struct _AddIn;

namespace addin
{
	EXTERN_C __declspec(
		dllexport) void __cdecl AddInLog(bool bPrintThreadTime, _Printf_format_string_ LPWSTR szMsg, ...);
	EXTERN_C __declspec(dllexport) void __cdecl GetMAPIModule(_In_ HMODULE* lphModule, bool bForce);

	void LoadAddIns();
	void UnloadAddIns();
	void MergeAddInArrays();
	std::wstring AddInSmartView(__ParsingTypeEnum iStructType, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	std::wstring AddInStructTypeToString(__ParsingTypeEnum iStructType);
} // namespace addin

extern std::vector<NAME_ARRAY_ENTRY_V2> PropTagArray;
extern std::vector<NAME_ARRAY_ENTRY> PropTypeArray;
extern std::vector<GUID_ARRAY_ENTRY> PropGuidArray;
extern std::vector<NAMEID_ARRAY_ENTRY> NameIDArray;
extern std::vector<FLAG_ARRAY_ENTRY> FlagArray;
extern std::vector<SMARTVIEW_PARSER_ARRAY_ENTRY> SmartViewParserArray;
extern std::vector<NAME_ARRAY_ENTRY> SmartViewParserTypeArray;
extern std::vector<_AddIn> g_lpMyAddins;