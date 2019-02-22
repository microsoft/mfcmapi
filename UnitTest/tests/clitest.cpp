#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/cli.h>

enum OPTIONFLAGS
{
	// Declared in flagsEnum
	//OPT_NOOPT = 0x0000,
	//OPT_INITMFC = 0x0001,
	//OPT_NEEDNUM = 0x0002, // Any arguments must be decimal numbers. No strings.
	OPT_10 = 0x0010,
	OPT_20 = 0x0020,
};

// Our set of options for testing
enum CmdMode
{
	cmdmode1 = cli::cmdmodeFirstMode,
	cmdmode2,
};

cli::option switchHelp = {L"?", cli::cmdmodeHelpFull, 0, 0, cli::OPT_INITMFC};
cli::option switchVerbose = {L"Verbose", cli::cmdmodeUnknown, 0, 0, cli::OPT_INITMFC};
cli::option switchMode1{L"mode1", cmdmode1, 0, 1, cli::OPT_NEEDNUM | OPT_10};
cli::option switchMode2{L"mode2", cmdmode2, 2, 2, OPT_20};

const std::vector<cli::option*> g_options = {&switchHelp, &switchVerbose, &switchMode1};

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
					expected->maxArgs != actual->maxArgs || expected->flags != actual->flags)
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
					Logger::WriteMessage(strings::format(L"ulOpt: %d:%d\n", expected->flags, actual->flags).c_str());
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
			AreEqual(&cli::switchHelp, GetOption(std::wstring{L"-?"}, g_options));
			AreEqual(&cli::switchVerbose, GetOption(std::wstring{L"-v"}, g_options));
			AreEqual(&cli::switchVerbose, GetOption(std::wstring{L"/v"}, g_options));
			AreEqual(&cli::switchVerbose, GetOption(std::wstring{L"\\v"}, g_options));
			AreEqual(&cli::switchVerbose, GetOption(std::wstring{L"-verbose"}, g_options));
			AreEqual(&cli::switchInvalid, GetOption(std::wstring{L"-verbosey"}, g_options));
			AreEqual(&cli::switchInvalid, GetOption(std::wstring{L"-va"}, g_options));
			AreEqual(&cli::switchInvalid, GetOption(std::wstring{L"-test"}, g_options));
			AreEqual(nullptr, GetOption(std::wstring{L""}, g_options));
			AreEqual(nullptr, GetOption(std::wstring{L"+v"}, g_options));
			AreEqual(&cli::switchInvalid, GetOption(std::wstring{L"-"}, g_options));

			AreEqual(nullptr, GetOption(L"No switch", g_options));
			AreEqual(&cli::switchInvalid, GetOption(L"-notaswitch", g_options));
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

		bool vetOption(
			const cli::option& option,
			const std::deque<std::wstring>& argsRemainder,
			const std::vector<std::wstring>& expectedArgs,
			const std::vector<std::wstring>& expectedRemainder)
		{
			if (option.size() != expectedArgs.size()) return false;
			for (UINT i = 0; i < option.size(); i++)
			{
				if (option.at(i) != expectedArgs[i]) return false;
			}

			if (argsRemainder.size() != expectedRemainder.size()) return false;
			for (UINT i = 0; i < argsRemainder.size(); i++)
			{
				if (argsRemainder[i] != expectedRemainder[i]) return false;
			}

			// In our failure case, we shouldn't have peeled off any args
			if (!option.isSet() && !option.empty()) return false;

			// After checking our args, if the option isn't set, we don't care about the rest
			if (!option.isSet()) return true;

			if (option.minArgs != 0 && option.empty()) return false;
			if (option.size() < option.minArgs) return false;
			if (option.size() > option.maxArgs) return false;

			if (option.flags & cli::OPT_NEEDNUM && !option.empty())
			{
				for (UINT i = 0; i < option.maxArgs; i++)
				{
					if (!option.hasULONG(i, 10)) return false;
				}
			}

			return true;
		}

		void test_scanArgs(
			bool targetResult,
			const cli::option& _option,
			const std::deque<std::wstring>& _args,
			const std::vector<std::wstring>& expectedArgs,
			const std::vector<std::wstring>& expectedRemainder)
		{
			// Make a local non-const copy of the inputs
			auto option = _option;
			auto args = _args;
			cli::OPTIONS options{};

			auto result = option.scanArgs(args, options, g_options);

			if (targetResult == result)
			{
				// If we claim to have matched, vet our options
				if (vetOption(option, args, expectedArgs, expectedRemainder))
				{
					return;
				}
			}

			Logger::WriteMessage(strings::format(L"scanArgs test failed\n").c_str());

			Logger::WriteMessage(strings::format(L"szSwitch: %ws\n", option.szSwitch).c_str());
			Logger::WriteMessage(strings::format(L"mode: %d\n", option.mode).c_str());
			Logger::WriteMessage(strings::format(L"minArgs: %d\n", option.minArgs).c_str());
			Logger::WriteMessage(strings::format(L"maxArgs: %d\n", option.maxArgs).c_str());
			Logger::WriteMessage(strings::format(L"ulOpt: %d\n", option.flags).c_str());
			Logger::WriteMessage(strings::format(L"seen: %d\n", option.isSet()).c_str());

			Logger::WriteMessage(strings::format(L"Tested args\n").c_str());
			for (auto& arg : args)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", arg.c_str()).c_str());
			}

			Logger::WriteMessage(strings::format(L"Parsed args\n").c_str());
			for (UINT i = 0; i < option.size(); i++)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", option.at(i).c_str()).c_str());
			}

			Logger::WriteMessage(strings::format(L"Expected args\n").c_str());
			for (auto& arg : expectedArgs)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", arg.c_str()).c_str());
			}

			Logger::WriteMessage(strings::format(L"Expected remainder\n").c_str());
			for (auto& arg : expectedRemainder)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", arg.c_str()).c_str());
			}

			Assert::Fail();
		}

		TEST_METHOD(Test_scanArgs)
		{
			// min/max-0/0
			auto p0_0 = cli::option{L"0_0", 0, 0, 0, 0};
			test_scanArgs(true, p0_0, {}, {}, {});
			test_scanArgs(true, p0_0, {L"-v"}, {}, {L"-v"});
			test_scanArgs(true, p0_0, {L"1"}, {}, {L"1"});

			// min/max-1/1
			auto p1_1 = cli::option{L"1_1", 0, 1, 1, 0};
			test_scanArgs(true, p1_1, {L"1"}, {L"1"}, {});
			test_scanArgs(true, p1_1, {L"2", L"3"}, {L"2"}, {L"3"});
			test_scanArgs(true, p1_1, {L"4", L"-v"}, {L"4"}, {L"-v"});
			// Not enough non switch args
			test_scanArgs(false, p1_1, {L"-v"}, {}, {L"-v"});
			// Not enough args at all
			test_scanArgs(false, p1_1, {}, {}, {});

			// min/max-0/1
			auto p0_1 = cli::option{L"0_1", 0, 0, 1, 0};
			test_scanArgs(true, p0_1, {L"-v"}, {}, {L"-v"});
			test_scanArgs(true, p0_1, {L"1", L"-v"}, {L"1"}, {L"-v"});
			test_scanArgs(true, p0_1, {}, {}, {});

			// min/max-2/3
			auto p2_3 = cli::option{L"2_3", 0, 2, 3, 0};
			test_scanArgs(true, p2_3, {L"1", L"2"}, {L"1", L"2"}, {});
			test_scanArgs(true, p2_3, {L"3", L"4", L"-v"}, {L"3", L"4"}, {L"-v"});
			test_scanArgs(true, p2_3, {L"5", L"6", L"7"}, {L"5", L"6", L"7"}, {});
			test_scanArgs(false, p2_3, {}, {}, {});
			test_scanArgs(false, p2_3, {L"1"}, {}, {L"1"});
			test_scanArgs(false, p2_3, {L"1", L"-v"}, {}, {L"1", L"-v"});

			auto p0_1_NUM = cli::option{L"0_1_NUM", 0, 0, 1, cli::OPT_NEEDNUM};
			test_scanArgs(true, p0_1_NUM, {L"1", L"2"}, {L"1"}, {L"2"});
			test_scanArgs(true, p0_1_NUM, {L"text", L"2"}, {}, {L"text", L"2"});
			test_scanArgs(true, p0_1_NUM, {}, {}, {});
		}
	}; // namespace clitest
} // namespace clitest