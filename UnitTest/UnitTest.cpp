#include <../StdAfx.h>
#include <CppUnitTest.h>
#include <UnitTest/UnitTest.h>
#include <UnitTest/resource.h>
#include <AddIns.h>
#include <IO/Registry.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace unittest
{
	void init()
	{
		// Set up our property arrays or nothing works
		addin::MergeAddInArrays();

		registry::doSmartView = true;
		registry::useGetPropList = true;
		registry::parseNamedProps = true;
		registry::cacheNamedProps = true;

		strings::setTestInstance(GetModuleHandleW(L"UnitTest.dll"));
	}

	// Assert::AreEqual doesn't do full logging, so we roll our own
	void AreEqualEx(
		const std::wstring& expected,
		const std::wstring& actual,
		const wchar_t* message,
		const __LineInfo* pLineInfo)
	{
		if (ignore_trailing_whitespace)
		{
			if (strings::trimWhitespace(expected) == strings::trimWhitespace(actual)) return;
		}
		else
		{
			if (expected == actual) return;
		}

		if (message != nullptr)
		{
			Logger::WriteMessage(strings::format(L"Test: %ws\n", message).c_str());
		}

		Logger::WriteMessage(L"Diff:\n");

		auto splitExpected = strings::split(expected, L'\n');
		auto splitActual = strings::split(actual, L'\n');
		auto errorCount = 0;
		for (size_t line = 0;
			 line < splitExpected.size() && line < splitActual.size() && (errorCount < 4 || !limit_output);
			 line++)
		{
			if (splitExpected[line] != splitActual[line])
			{
				errorCount++;
				Logger::WriteMessage(
					strings::format(
						L"[%d]\n\"%ws\"\n\"%ws\"\n", line + 1, splitExpected[line].c_str(), splitActual[line].c_str())
						.c_str());
				auto lineErrorCount = 0;
				for (size_t ch = 0; ch < splitExpected[line].size() && ch < splitActual[line].size() &&
									(lineErrorCount < 10 || !limit_output);
					 ch++)
				{
					const auto expectedChar = splitExpected[line][ch];
					const auto actualChar = splitActual[line][ch];
					if (expectedChar != actualChar)
					{
						lineErrorCount++;
						Logger::WriteMessage(strings::format(
												 L"[%d]: %X (%wc) != %X (%wc)\n",
												 ch + 1,
												 expectedChar,
												 expectedChar,
												 actualChar,
												 actualChar)
												 .c_str());
					}
				}
			}
		}

		Logger::WriteMessage(L"\n");
		Logger::WriteMessage(strings::format(L"Expected:\n\"%ws\"\n\n", expected.c_str()).c_str());
		Logger::WriteMessage(strings::format(L"Actual:\n\"%ws\"", actual.c_str()).c_str());

		if (assert_on_failure)
		{
			Assert::Fail(ToString(message).c_str(), pLineInfo);
		}
	}

	// Resource files saved in unicode have a byte order mark of 0xfffe
	// We load these in and strip the BOM.
	// Otherwise we load as ansi and convert to unicode
	std::wstring loadfile(const HMODULE handle, const int name)
	{
		const auto rc = ::FindResource(handle, MAKEINTRESOURCE(name), MAKEINTRESOURCE(TEXTFILE));
		const auto rcData = LoadResource(handle, rc);
		const auto cb = SizeofResource(handle, rc);
		const auto bytes = LockResource(rcData);
		const auto data = static_cast<const BYTE*>(bytes);

		// UTF 16 LE
		if (cb >= 2 && data[0] == 0xff && data[1] == 0xfe)
		{
			// Skip the byte order mark
			const auto wstr = static_cast<const wchar_t*>(bytes);
			const auto cch = cb / sizeof(wchar_t);
			return std::wstring(wstr + 1, cch - 1);
		}

		const auto str = std::string(static_cast<const char*>(bytes), cb);
		return strings::stringTowstring(str);
	}
} // namespace unittest