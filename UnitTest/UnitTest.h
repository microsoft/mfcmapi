#pragma once
#define VC_EXTRALEAN // Exclude rarely-used stuff from Windows headers

#define _WINSOCKAPI_ // stops windows.h including winsock.h
#include <Windows.h>
#include <tchar.h>

// Common CRT headers
#include <list>
#include <algorithm>
#include <locale>
#include <sstream>
#include <iterator>
#include <functional>
#include <map>
#include <unordered_map>
#include <utility>
#include <vector>
#include <cassert>
#include <stack>

// All the MAPI headers
#include <MAPIX.h>
#include <MAPIUtil.h>
#include <MAPIForm.h>
#include <MAPIWz.h>
#include <MAPIHook.h>
#include <MSPST.h>

#include <EdkMdb.h>
#include <ExchForm.h>
#include <EMSABTAG.H>
#include <IMessage.h>
#include <EdkGuid.h>
#include <TNEF.h>
#include <MAPIAux.h>

// For import procs
#include <AclUI.h>
#include <Uxtheme.h>

// there's an odd conflict with mimeole.h and richedit.h - this should fix it
#ifdef UNICODE
#undef CHARFORMAT
#endif
#include <mimeole.h>
#ifdef UNICODE
#undef CHARFORMAT
#define CHARFORMAT CHARFORMATW
#endif

#include <core/res/Resource.h> // main symbols
#include <CppUnitTest.h>
#include <core/interpret/guid.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template <> inline std::wstring ToString<__int64>(const __int64& t) { RETURN_WIDE_STRING(t); }
			template <> inline std::wstring ToString<std::vector<BYTE>>(const std::vector<BYTE>& t)
			{
				RETURN_WIDE_STRING(t.data());
			}
			template <> inline std::wstring ToString<GUID>(const GUID& t) { return guid::GUIDToString(t); }
		} // namespace CppUnitTestFramework
	} // namespace VisualStudio
} // namespace Microsoft

namespace unittest
{
	static const bool parse_all = true;
	static const bool assert_on_failure = true;
	static const bool limit_output = true;
	static const bool ignore_trailing_whitespace = false;

	void init();

	void AreEqualEx(
		const std::wstring& expected,
		const std::wstring& actual,
		const wchar_t* message = nullptr,
		const __LineInfo* pLineInfo = nullptr);

	std::wstring loadfile(HMODULE handle, int name);
} // namespace unittest