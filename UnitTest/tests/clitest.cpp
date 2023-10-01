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
	cmdmode3,
	cmdmode4,
};

#pragma warning(push)
#pragma warning(disable : 5054) // Warning C5054 operator '|': deprecated between enumerations of different types
cli::option switchMode1{L"mode1", cmdmode1, 0, 1, cli::OPT_NEEDNUM | OPT_10};
cli::option switchMode2{L"mode2", cmdmode2, 2, 2, OPT_20};
#pragma warning(pop)

const std::vector<cli::option*> g_options = {&cli::switchHelp, &cli::switchVerbose, &switchMode1, &switchMode2};

namespace Microsoft::VisualStudio::CppUnitTestFramework
{
	template <> inline std::wstring ToString<cli::modeEnum>(const cli::modeEnum& q) { RETURN_WIDE_STRING(q); }

	void AreEqual(
		const cli::option* expected,
		const cli::option* actual,
		const wchar_t* message = nullptr,
		const __LineInfo* pLineInfo = nullptr)
	{
		if (expected == actual) return;

		Logger::WriteMessage(strings::format(L"Switch:\n").c_str());
		if (expected) Logger::WriteMessage(strings::format(L"Expected: %ws\n", expected->name()).c_str());
		if (actual) Logger::WriteMessage(strings::format(L"Actual: %ws\n", actual->name()).c_str());

		Assert::Fail(ToString(message).c_str(), pLineInfo);
	}
} // namespace Microsoft::VisualStudio::CppUnitTestFramework

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
			auto noarg = std::vector<LPCWSTR>{L"app.exe"};
			Assert::AreEqual({}, cli::GetCommandLine(static_cast<int>(noarg.size()), noarg.data()));
			auto argv = std::vector<LPCWSTR>{L"app.exe", L"-arg1", L"-arg2", L"testü", L"ěřů"};
			Assert::AreEqual(
				{L"-arg1", L"-arg2", L"testü", L"ěřů"},
				cli::GetCommandLine(static_cast<int>(argv.size()), argv.data()));
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
			const std::vector<std::wstring>& expectedRemainder) noexcept
		{
			if (option.size() != expectedArgs.size()) return false;
			for (UINT i = 0; i < option.size(); i++)
			{
				if (option[i] != expectedArgs[i]) return false;
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
				for (UINT i = 0; i < option.size(); i++)
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

			const auto result = option.scanArgs(args, options, g_options);

			if (targetResult == result)
			{
				// If we claim to have matched, vet our options
				if (vetOption(option, args, expectedArgs, expectedRemainder))
				{
					return;
				}
			}

			Logger::WriteMessage(strings::format(L"scanArgs test failed\n").c_str());

			Logger::WriteMessage(strings::format(L"szSwitch: %ws\n", option.name()).c_str());
			Logger::WriteMessage(strings::format(L"mode: %d\n", option.mode).c_str());
			Logger::WriteMessage(strings::format(L"minArgs: %d\n", option.minArgs).c_str());
			Logger::WriteMessage(strings::format(L"maxArgs: %d\n", option.maxArgs).c_str());
			Logger::WriteMessage(strings::format(L"ulOpt: %d\n", option.flags).c_str());
			Logger::WriteMessage(strings::format(L"seen: %d\n", option.isSet()).c_str());

			Logger::WriteMessage(strings::format(L"Tested args\n").c_str());
			for (auto& arg : _args)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", arg.c_str()).c_str());
			}

			Logger::WriteMessage(strings::format(L"Parsed args\n").c_str());
			for (UINT i = 0; i < option.size(); i++)
			{
				Logger::WriteMessage(strings::format(L"  %ws\n", option[i].c_str()).c_str());
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

			auto p1_2_NUM = cli::option{L"1_2_NUM", 0, 1, 2, cli::OPT_NEEDNUM};
			test_scanArgs(true, p1_2_NUM, {L"1", L"2", L"3"}, {L"1", L"2"}, {L"3"});
			test_scanArgs(true, p1_2_NUM, {L"1", L"text", L"3"}, {L"1"}, {L"text", L"3"});
			test_scanArgs(false, p1_2_NUM, {L"text", L"2", L"3"}, {}, {L"text", L"2", L"3"});
			test_scanArgs(false, p1_2_NUM, {L"text", L"2"}, {}, {L"text", L"2"});
			test_scanArgs(false, p1_2_NUM, {}, {}, {});
		}

		TEST_METHOD(Test_option)
		{
			cli::option switch1{L"switch1", cli::cmdmodeHelpFull, 2, 2, 0};
			const std::vector<cli::option*> optionsArray = {&switch1};
			cli::OPTIONS options{};

			Assert::AreEqual(L"switch1", switch1.name());
			Assert::AreEqual(std::wstring{L""}, switch1.at(1));
			Assert::AreEqual(false, switch1.isSet());
			Assert::AreEqual(true, switch1.empty());
			Assert::AreEqual(std::wstring{L""}, switch1.at(0));
			Assert::AreEqual(std::wstring{L""}, switch1[0]);
			Assert::AreEqual(false, switch1.has(0));
			Assert::AreEqual(false, switch1.hasULONG(0));
			Assert::AreEqual(false, switch1.hasULONG(1));
			Assert::AreEqual(ULONG{0}, switch1.atULONG(0));
			Assert::AreEqual(ULONG{0}, switch1.atULONG(1));
			Assert::AreEqual(size_t{0}, switch1.size());

			auto args = std::deque<std::wstring>{L"3"};
			Assert::AreEqual(false, switch1.scanArgs(args, options, optionsArray));
			Assert::AreEqual(size_t{0}, switch1.size());
			args.push_back({L"text"});
			Assert::AreEqual(true, switch1.scanArgs(args, options, optionsArray));
			Assert::AreEqual(size_t{2}, switch1.size());
			Assert::AreEqual(true, switch1.isSet());
			Assert::AreEqual(false, switch1.empty());
			Assert::AreEqual(std::wstring{L"3"}, switch1.at(0));
			Assert::AreEqual(std::wstring{L"3"}, switch1[0]);
			Assert::AreEqual(true, switch1.has(0));
			Assert::AreEqual(true, switch1.hasULONG(0));
			Assert::AreEqual(false, switch1.hasULONG(1));
			Assert::AreEqual(ULONG{3}, switch1.atULONG(0));
			Assert::AreEqual(ULONG{0}, switch1.atULONG(1));
		}

		void test_ParseArgs(
			const std::deque<std::wstring>& _args,
			const std::vector<cli::option*>& optionsArray,
			const int mode,
			const std::vector<cli::option*>& setOptions,
			const std::wstring& unswitchedArg)
		{
			cli::OPTIONS options{};
			auto args = _args;
			cli::ParseArgs(options, args, optionsArray);
			auto eq = true;

			if (options.mode != mode) eq = false;
			size_t i{};
			for (i = 0; i < setOptions.size(); i++)
			{
				if (!setOptions[i]->isSet()) eq = false;
				if (setOptions[i]->size() < setOptions[i]->minArgs) eq = false;
				if (setOptions[i]->size() > setOptions[i]->maxArgs) eq = false;
			}

			if (!eq)
			{
				Logger::WriteMessage(strings::format(L"Expected mode: %i\n", mode).c_str());
				Logger::WriteMessage(strings::format(L"Actual mode: %i\n", options.mode).c_str());
				Logger::WriteMessage(
					strings::format(L"Expected unswitched option: %ws\n", unswitchedArg.c_str()).c_str());
				Logger::WriteMessage(
					strings::format(L"Actual unswitched option: %ws\n", cli::switchUnswitched[0].c_str()).c_str());

				Logger::WriteMessage(strings::format(L"Set option count: %i\n", setOptions.size()).c_str());
				for (i = 0; i < setOptions.size(); i++)
				{
					Logger::WriteMessage(
						strings::format(L"[%i] = %ws = %i\n", i, setOptions[i]->name(), setOptions[i]->isSet())
							.c_str());
				}

				Assert::Fail();
			}
		}

		TEST_METHOD(Test_ParseArgs)
		{
			cli::option switch1{L"switch1", cmdmode1, 2, 2, 0};
			cli::option switch2{L"switch2", cmdmode2, 0, 0, 0};
			test_ParseArgs({}, {&switch1, &switch2}, cli::cmdmodeHelp, {}, {});
			test_ParseArgs({L"-switch1", L"3", L"4"}, {&switch1, &switch2}, cmdmode1, {&switch1}, {});
			test_ParseArgs(
				{L"-switch1", L"3", L"4", L"-switch2"}, {&switch1, &switch2}, cli::cmdmodeHelp, {&switch1}, {});
			test_ParseArgs({L"-notaswitch", L"3", L"4"}, {&switch1, &switch2}, cli::cmdmodeHelp, {}, {});
			test_ParseArgs({L"unswitch", L"-switch2"}, {&switch1, &switch2}, cmdmode2, {&switch2}, L"unswitch");
			test_ParseArgs({L"unswitch1", L"unswitch2"}, {&switch1, &switch2}, cli::cmdmodeHelp, {}, {});
		}
	};
} // namespace clitest