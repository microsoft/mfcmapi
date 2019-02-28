#pragma once
#include <CppUnitTest.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template <> inline std::wstring ToString<__int64>(const __int64& q) { RETURN_WIDE_STRING(q); }
			template <> inline std::wstring ToString<std::vector<BYTE>>(const std::vector<BYTE>& q)
			{
				RETURN_WIDE_STRING(q.data());
			}
			template <> inline std::wstring ToString<GUID>(const GUID& q) { return guid::GUIDToString(q); }
			template <> inline std::wstring ToString<std::deque<std::wstring>>(const std::deque<std::wstring>& q)
			{
				return strings::join({q.begin(), q.end()}, L",");
			}
		} // namespace CppUnitTestFramework
	} // namespace VisualStudio
} // namespace Microsoft

namespace unittest
{
	static const bool parse_all = false;
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