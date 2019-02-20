#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/cli.h>

// Our set of options for testing
const std::vector<cli::option*> options = {&cli::switchHelp, &cli::switchVerbose};

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template <> inline std::wstring ToString<cli::modeEnum>(const cli::modeEnum& q) { RETURN_WIDE_STRING(q); }

			void AreEqual(
				const cli::option* expected,
				const cli::option* actual,
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

		TEST_METHOD(Test_GetOption)
		{
			AreEqual(&cli::switchHelp, GetOption(std::wstring{L"-?"}, options));
			AreEqual(&cli::switchVerbose, GetOption(std::wstring{L"-v"}, options));
			AreEqual(&cli::switchVerbose, GetOption(std::wstring{L"/v"}, options));
			AreEqual(&cli::switchVerbose, GetOption(std::wstring{L"\\v"}, options));
			AreEqual(&cli::switchVerbose, GetOption(std::wstring{L"-verbose"}, options));
			AreEqual(&cli::switchInvalid, GetOption(std::wstring{L"-verbosey"}, options));
			AreEqual(&cli::switchInvalid, GetOption(std::wstring{L"-va"}, options));
			AreEqual(&cli::switchInvalid, GetOption(std::wstring{L"-test"}, options));
			AreEqual(nullptr, GetOption(std::wstring{L""}, options));
			AreEqual(nullptr, GetOption(std::wstring{L"+v"}, options));
			AreEqual(&cli::switchInvalid, GetOption(std::wstring{L"-"}, options));

			AreEqual(nullptr, GetOption(L"No switch", options));
			AreEqual(&cli::switchInvalid, GetOption(L"-notaswitch", options));
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

		void test_scanArgs(bool targetResult, const cli::option& _option, const std::deque<std::wstring>& _args)
		{
			// Make a local non-const copy of the inputs
			auto option = _option;
			auto args = _args;

			auto result = option.scanArgs(args, options);

			if (targetResult == result)
			{
				if (!result) return;

				// If we claim to have matched, some further checks
				if (option.args.size() >= option.minArgs && option.args.size() <= option.maxArgs)
				{
					return;
				}
			}

			Logger::WriteMessage(strings::format(L"scanArgs test failed\n").c_str());

			Logger::WriteMessage(strings::format(L"Switch: %ws\n", option.szSwitch).c_str());
			Logger::WriteMessage(strings::format(L"Mode: %d\n", option.mode).c_str());
			Logger::WriteMessage(strings::format(L"minArgs: %d\n", option.minArgs).c_str());
			Logger::WriteMessage(strings::format(L"maxArgs: %d\n", option.maxArgs).c_str());
			Logger::WriteMessage(strings::format(L"ulOpt: %d\n", option.options).c_str());

			Logger::WriteMessage(strings::format(L"Tested args\n").c_str());
			for (auto& arg : args)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", arg.c_str()).c_str());
			}

			Logger::WriteMessage(strings::format(L"Parsed args\n").c_str());
			for (auto& arg : option.args)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", arg.c_str()).c_str());
			}

			Assert::Fail();
		}

		TEST_METHOD(Test_scanArgs)
		{
			// min/max-0/0
			auto p0_0 = cli::option{L"", 0, 0, 0, 0};
			test_scanArgs(true, p0_0, {});
			test_scanArgs(true, p0_0, {});
			test_scanArgs(true, p0_0, {L"-v"});
			test_scanArgs(true, p0_0, {L"1"});

			//// min/max-1/1
			auto p1_1 = cli::option{L"", 0, 1, 1, 0};
			test_scanArgs(true, p1_1, {L"1"});
			test_scanArgs(true, p1_1, {L"2", L"3"});
			test_scanArgs(true, p1_1, {L"4", L"-v"});
			// Not enough non switch args
			test_scanArgs(false, p1_1, {L"-v"});
			// Not enough args at all
			test_scanArgs(false, p1_1, {});

			// min/max-0/1
			auto p0_1 = cli::option{L"", 0, 0, 1, 0};
			test_scanArgs(true, p0_1, {L"-v"});
			test_scanArgs(true, p0_1, {L"1", L"-v"});
			test_scanArgs(true, p0_1, {});

			// min/max-2/3
			auto p2_3 = cli::option{L"", 0, 2, 3, 0};
			test_scanArgs(true, p2_3, {L"1", L"2"});
			test_scanArgs(true, p2_3, {L"3", L"4", L"-v"});
			test_scanArgs(true, p2_3, {L"5", L"6", L"7"});
			test_scanArgs(false, p2_3, {});
			test_scanArgs(false, p2_3, {L"1"});
			test_scanArgs(false, p2_3, {L"1", L"-v"});
		}
	}; // namespace clitest
} // namespace clitest