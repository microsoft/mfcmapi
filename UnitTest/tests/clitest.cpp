#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/cli.h>

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template <> inline std::wstring ToString<cli::modeEnum>(const cli::modeEnum& q) { RETURN_WIDE_STRING(q); }

			void AreEqual(
				const cli::OptParser* expected,
				const cli::OptParser* actual,
				const wchar_t* message = nullptr,
				const __LineInfo* pLineInfo = nullptr)
			{
				if (expected == nullptr && actual == nullptr) return;
				if (expected == nullptr || actual == nullptr)
				{
					Logger::WriteMessage(strings::format(L"expected: %p\n", expected).c_str());
					Logger::WriteMessage(strings::format(L"actual: %p\n", actual).c_str());
				}

				auto eq = true;
				if (std::wstring{expected->szSwitch} != std::wstring{actual->szSwitch} ||
					expected->mode != actual->mode || expected->minArgs != actual->minArgs ||
					expected->maxArgs != actual->maxArgs || expected->options != actual->options)
				{
					eq = false;
				}

				if (!eq)
				{
					Logger::WriteMessage(
						strings::format(L"Switch: %ws:%ws\n", expected->szSwitch, actual->szSwitch).c_str());
					Logger::WriteMessage(strings::format(L"Mode: %d:%d\n", expected->mode, actual->mode).c_str());
					Logger::WriteMessage(
						strings::format(L"minArgs: %d:%d\n", expected->minArgs, actual->minArgs).c_str());
					Logger::WriteMessage(
						strings::format(L"maxArgs: %d:%d\n", expected->maxArgs, actual->maxArgs).c_str());
					Logger::WriteMessage(
						strings::format(L"ulOpt: %d:%d\n", expected->options, actual->options).c_str());
					Assert::Fail(ToString(message).c_str(), pLineInfo);
				}
			}

		} // namespace CppUnitTestFramework
	} // namespace VisualStudio
} // namespace Microsoft

namespace clitest
{
	TEST_CLASS(clitest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_GetCommandLine)
		{
			//	std::vector<std::wstring> GetCommandLine(_In_ int argc, _In_count_(argc) const char* const argv[]);
			auto noarg = std::vector<LPCSTR>{"app.exe"};
			Assert::AreEqual({}, cli::GetCommandLine(int(noarg.size()), noarg.data()));
			auto argv = std::vector<LPCSTR>{"app.exe", "-arg1", "-arg2"};
			Assert::AreEqual({L"-arg1", L"-arg2"}, cli::GetCommandLine(int(argv.size()), argv.data()));
		}

		TEST_METHOD(Test_GetParser)
		{
			AreEqual(&cli::switchHelpParser, GetParser(std::wstring{L"-?"}, cli::parsers));
			AreEqual(&cli::switchVerboseParser, GetParser(std::wstring{L"-v"}, cli::parsers));
			AreEqual(&cli::switchVerboseParser, GetParser(std::wstring{L"/v"}, cli::parsers));
			AreEqual(&cli::switchVerboseParser, GetParser(std::wstring{L"\\v"}, cli::parsers));
			AreEqual(&cli::switchVerboseParser, GetParser(std::wstring{L"-verbose"}, cli::parsers));
			AreEqual(&cli::switchInvalidParser, GetParser(std::wstring{L"-verbosey"}, cli::parsers));
			AreEqual(&cli::switchInvalidParser, GetParser(std::wstring{L"-va"}, cli::parsers));
			AreEqual(&cli::switchInvalidParser, GetParser(std::wstring{L"-test"}, cli::parsers));
			AreEqual(nullptr, GetParser(std::wstring{L""}, cli::parsers));
			AreEqual(nullptr, GetParser(std::wstring{L"+v"}, cli::parsers));
			AreEqual(&cli::switchInvalidParser, GetParser(std::wstring{L"-"}, cli::parsers));

			AreEqual(nullptr, GetParser(L"No switch", cli::parsers));
			AreEqual(&cli::switchInvalidParser, GetParser(L"-notaswitch", cli::parsers));
		}

		TEST_METHOD(Test_bSetMode)
		{
			auto mode = int{};
			Assert::AreEqual(true, cli::bSetMode(mode, cli::cmdmodeHelpFull));
			Assert::AreEqual(true, cli::bSetMode(mode, cli::cmdmodeHelpFull));
			Assert::AreEqual(cli::cmdmodeHelpFull, cli::modeEnum(mode));
			Assert::AreEqual(false, cli::bSetMode(mode, cli::cmdmodeHelp));
			Assert::AreEqual(cli::cmdmodeHelpFull, cli::modeEnum(mode));
		}

		void test_scanArgs(bool targetResult, const cli::OptParser& _parser, const std::deque<std::wstring>& _args)
		{
			// Make a local non-const copy of the inputs
			auto parser = _parser;
			auto args = _args;

			auto result = parser.scanArgs(args, cli::parsers);

			if (targetResult == result)
			{
				if (!result) return;

				// If we claim to have matched, some further checks
				if (parser.args.size() >= parser.minArgs && parser.args.size() <= parser.maxArgs)
				{
					return;
				}
			}

			Logger::WriteMessage(strings::format(L"scanArgs test failed\n").c_str());

			Logger::WriteMessage(strings::format(L"Switch: %ws\n", parser.szSwitch).c_str());
			Logger::WriteMessage(strings::format(L"Mode: %d\n", parser.mode).c_str());
			Logger::WriteMessage(strings::format(L"minArgs: %d\n", parser.minArgs).c_str());
			Logger::WriteMessage(strings::format(L"maxArgs: %d\n", parser.maxArgs).c_str());
			Logger::WriteMessage(strings::format(L"ulOpt: %d\n", parser.options).c_str());

			Logger::WriteMessage(strings::format(L"Tested args\n").c_str());
			for (auto& arg : args)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", arg.c_str()).c_str());
			}

			Logger::WriteMessage(strings::format(L"Parsed args\n").c_str());
			for (auto& arg : parser.args)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", arg.c_str()).c_str());
			}

			Assert::Fail();
		}

		TEST_METHOD(Test_scanArgs)
		{
			// min/max-0/0
			auto p0_0 = cli::OptParser{L"", 0, 0, 0, 0};
			test_scanArgs(true, p0_0, {});
			test_scanArgs(true, p0_0, {});
			test_scanArgs(true, p0_0, {L"-v"});
			test_scanArgs(true, p0_0, {L"1"});

			//// min/max-1/1
			auto p1_1 = cli::OptParser{L"", 0, 1, 1, 0};
			test_scanArgs(true, p1_1, {L"1"});
			test_scanArgs(true, p1_1, {L"2", L"3"});
			test_scanArgs(true, p1_1, {L"4", L"-v"});
			// Not enough non switch args
			test_scanArgs(false, p1_1, {L"-v"});
			// Not enough args at all
			test_scanArgs(false, p1_1, {});

			// min/max-0/1
			auto p0_1 = cli::OptParser{L"", 0, 0, 1, 0};
			test_scanArgs(true, p0_1, {L"-v"});
			test_scanArgs(true, p0_1, {L"1", L"-v"});
			test_scanArgs(true, p0_1, {});

			// min/max-2/3
			auto p2_3 = cli::OptParser{L"", 0, 2, 3, 0};
			test_scanArgs(true, p2_3, {L"1", L"2"});
			test_scanArgs(true, p2_3, {L"3", L"4", L"-v"});
			test_scanArgs(true, p2_3, {L"5", L"6", L"7"});
			test_scanArgs(false, p2_3, {});
			test_scanArgs(false, p2_3, {L"1"});
			test_scanArgs(false, p2_3, {L"1", L"-v"});
		}
	}; // namespace clitest
} // namespace clitest