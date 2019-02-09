#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/cli.h>

enum switchEnum
{
	switchNoSwitch = 0, // not a switch
	switchUnknown, // unknown switch
	switchHelp, // '-h'
	switchVerbose, // '-v'
	switchSearch, // '-s'
};

enum flagsEnum
{
	OPT_NOOPT = 0x00000,
	OPT_DOPARTIALSEARCH = 0x00001,
	OPT_VERBOSE = 0x00200,
	OPT_INITMFC = 0x10000,
};
flagsEnum& operator|=(flagsEnum& a, flagsEnum b)
{
	return reinterpret_cast<flagsEnum&>(reinterpret_cast<int&>(a) |= static_cast<int>(b));
}
flagsEnum operator|(flagsEnum a, flagsEnum b)
{
	return static_cast<flagsEnum>(static_cast<int>(a) | static_cast<int>(b));
}

enum modeEnum
{
	cmdmodeUnknown = 0,
	cmdmodeHelp,
	cmdmodeHelpFull,
	cmdmodePropTag,
	cmdmodeGuid,
};

const std::vector<cli::COMMANDLINE_SWITCH> switches = {
	{switchNoSwitch, L""},
	{switchUnknown, L""},
	{switchHelp, L"?"},
	{switchVerbose, L"Verbose"},
	{switchSearch, L"Search"},
};

const std::vector<cli::OptParser> parsers = {
	{switchHelp, cmdmodeHelpFull, 0, 0, OPT_INITMFC},
	{switchVerbose, cmdmodeUnknown, 0, 0, OPT_VERBOSE | OPT_INITMFC},
	{switchSearch, cmdmodeUnknown, 0, 0, OPT_DOPARTIALSEARCH},
	{switchNoSwitch, cmdmodeUnknown, 0, 0, OPT_NOOPT},
};

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			template <> inline std::wstring ToString<modeEnum>(const modeEnum& q) { RETURN_WIDE_STRING(q); }

			void AreEqual(
				const cli::OptParser& expected,
				const cli::OptParser& actual,
				const wchar_t* message = nullptr,
				const __LineInfo* pLineInfo = nullptr)
			{
				auto eq = true;
				if (expected.clSwitch != actual.clSwitch || expected.mode != actual.mode ||
					expected.minArgs != actual.minArgs || expected.maxArgs != actual.maxArgs ||
					expected.options != actual.options)
				{
					eq = false;
				}

				if (!eq)
				{
					Logger::WriteMessage(
						strings::format(L"Switch: %d:%d\n", expected.clSwitch, actual.clSwitch).c_str());
					Logger::WriteMessage(strings::format(L"Mode: %d:%d\n", expected.mode, actual.mode).c_str());
					Logger::WriteMessage(
						strings::format(L"minArgs: %d:%d\n", expected.minArgs, actual.minArgs).c_str());
					Logger::WriteMessage(
						strings::format(L"maxArgs: %d:%d\n", expected.maxArgs, actual.maxArgs).c_str());
					Logger::WriteMessage(strings::format(L"ulOpt: %d:%d\n", expected.options, actual.options).c_str());
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
			Assert::AreEqual(std::deque<std::wstring>{}, cli::GetCommandLine(int(noarg.size()), noarg.data()));
			auto argv = std::vector<LPCSTR>{"app.exe", "-arg1", "-arg2"};
			Assert::AreEqual(
				std::deque<std::wstring>{L"-arg1", L"-arg2"}, cli::GetCommandLine(int(argv.size()), argv.data()));
		}

		TEST_METHOD(Test_ParseArgument)
		{
			Assert::AreEqual(int(switchHelp), int(ParseArgument(std::wstring{L"-?"}, switches)));
			Assert::AreEqual(int(switchVerbose), int(ParseArgument(std::wstring{L"-v"}, switches)));
			Assert::AreEqual(int(switchVerbose), int(ParseArgument(std::wstring{L"/v"}, switches)));
			Assert::AreEqual(int(switchVerbose), int(ParseArgument(std::wstring{L"\\v"}, switches)));
			Assert::AreEqual(int(switchVerbose), int(ParseArgument(std::wstring{L"-verbose"}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(ParseArgument(std::wstring{L"-verbosey"}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(ParseArgument(std::wstring{L"-va"}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(ParseArgument(std::wstring{L"-test"}, switches)));
			Assert::AreEqual(int(switchSearch), int(ParseArgument(std::wstring{L"/s"}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(ParseArgument(std::wstring{L""}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(ParseArgument(std::wstring{L"+v"}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(ParseArgument(std::wstring{L"-"}, switches)));
		}

		TEST_METHOD(Test_GetParser)
		{
			AreEqual(parsers[0], GetParser(switchHelp, parsers));
			AreEqual(parsers[1], GetParser(switchVerbose, parsers));
			AreEqual({}, GetParser(switchNoSwitch, parsers));
		}

		TEST_METHOD(Test_bSetMode)
		{
			auto mode = int{};
			Assert::AreEqual(true, cli::bSetMode(mode, cmdmodeHelpFull));
			Assert::AreEqual(true, cli::bSetMode(mode, cmdmodeHelpFull));
			Assert::AreEqual(cmdmodeHelpFull, modeEnum(mode));
			Assert::AreEqual(false, cli::bSetMode(mode, cmdmodePropTag));
			Assert::AreEqual(cmdmodeHelpFull, modeEnum(mode));
		}
	}; // namespace clitest
} // namespace clitest