#pragma once
#include <StdAfx.h>
#include <CppUnitTest.h>
#include <Interpret/SmartView/SmartView.h>
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