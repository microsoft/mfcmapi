#pragma once
#include <CppUnitTest.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/addin/mfcmapi.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace Microsoft::VisualStudio::CppUnitTestFramework
{
	template <> inline std::wstring ToString<std::vector<BYTE>>(const std::vector<BYTE>& q)
	{
		RETURN_WIDE_STRING(q.data());
	}
	template <> inline std::wstring ToString<GUID>(const GUID& q) { return guid::GUIDToString(q); }
	template <> inline std::wstring ToString<std::deque<std::wstring>>(const std::deque<std::wstring>& q)
	{
		return strings::join({q.begin(), q.end()}, L",");
	}
	template <> inline std::wstring ToString<std::vector<std::wstring>>(const std::vector<std::wstring>& q)
	{
		return strings::join({q.begin(), q.end()}, L",");
	}
} // namespace Microsoft::VisualStudio::CppUnitTestFramework

namespace unittest
{
	static constexpr bool parse_all = true;
	static constexpr bool assert_on_failure = true;
	static constexpr bool limit_output = true;
	static constexpr bool ignore_trailing_whitespace = false;

	void init();

	void AreEqualEx(
		const std::wstring& expected,
		const std::wstring& actual,
		const wchar_t* message = nullptr,
		const __LineInfo* pLineInfo = nullptr);

	std::wstring loadfile(HMODULE handle, int name);

	void test(const std::wstring testName, parserType structType, std::vector<BYTE> hex, const std::wstring expected);
	void test(parserType structType, DWORD hexNum, DWORD expectedNum);
} // namespace unittest