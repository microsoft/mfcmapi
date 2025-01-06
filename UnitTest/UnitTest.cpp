#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/smartview/SmartView.h>

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

		Logger::WriteMessage(L"Diff (expected vs actual):\n");

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
		if (!rc) return strings::emptystring;
		const auto rcData = LoadResource(handle, rc);
		if (!rcData) return strings::emptystring;
		const auto cb = SizeofResource(handle, rc);
		const auto bytes = LockResource(rcData);
		const auto data = static_cast<const BYTE*>(bytes);

		// UTF 16 LE
		// In Notepad++, this is UCS-2 LE BOM encoding
		// WARNING: Editing files in Visual Studio Code can alter this encoding
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

	void test(const std::wstring testName, parserType structType, std::vector<BYTE> hex, const std::wstring expected)
	{
		auto actual =
			smartview::InterpretBinary({static_cast<ULONG>(hex.size()), hex.data()}, structType, nullptr)->toString();
		unittest::AreEqualEx(expected, actual, testName.c_str());

		if (unittest::parse_all)
		{
			for (const auto parser : SmartViewParserTypeArray)
			{
				if (parser.type == parserType::NOPARSING) continue;
				try
				{
					actual =
						smartview::InterpretBinary({static_cast<ULONG>(hex.size()), hex.data()}, parser.type, nullptr)
							->toString();
				} catch (const int exception)
				{
					Logger::WriteMessage(strings::format(
											 L"Testing %ws failed at %ws with error 0x%08X\n",
											 testName.c_str(),
											 addin::AddInStructTypeToString(parser.type).c_str(),
											 exception)
											 .c_str());
					Assert::Fail();
				}
			}
		}
	}

	void test(parserType structType, DWORD hexNum, DWORD expectedNum)
	{
		static auto handle = GetModuleHandleW(L"UnitTest.dll");
		// See comments on loadfile for best file encoding strategies for test data
		const auto testName = strings::format(L"%d/%d", hexNum, expectedNum);
		auto hex = strings::HexStringToBin(unittest::loadfile(handle, hexNum));
		const auto expected = unittest::loadfile(handle, expectedNum);
		test(testName, structType, hex, expected);
	}
} // namespace unittest