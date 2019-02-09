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

cli::OptParser noSwitchParser = {switchNoSwitch, cmdmodeUnknown, 0, 0, OPT_NOOPT};
cli::OptParser helpParser = {switchHelp, cmdmodeHelpFull, 0, 0, OPT_INITMFC};
cli::OptParser verboseParser = {switchVerbose, cmdmodeUnknown, 0, 0, OPT_VERBOSE | OPT_INITMFC};
cli::OptParser searchParser = {switchSearch, cmdmodeUnknown, 0, 0, OPT_DOPARTIALSEARCH};

const std::vector<cli::OptParser> parsers = {noSwitchParser, helpParser, verboseParser, searchParser};

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
			Assert::AreEqual({}, cli::GetCommandLine(int(noarg.size()), noarg.data()));
			auto argv = std::vector<LPCSTR>{"app.exe", "-arg1", "-arg2"};
			Assert::AreEqual({L"-arg1", L"-arg2"}, cli::GetCommandLine(int(argv.size()), argv.data()));
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
			AreEqual(helpParser, GetParser(switchHelp, parsers));
			AreEqual(verboseParser, GetParser(switchVerbose, parsers));
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

		TEST_METHOD(Test_CheckMinArgs)
		{
			// min/max-0/0
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 0, 0, 0}, {L"-v"}, switches));
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 0, 0, 0}, {L"-v", L"-v"}, switches));
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 0, 0, 0}, {L"-v", L"1"}, switches));

			// min/max-1/1
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 1, 1, 0}, {L"-v", L"1"}, switches));
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 1, 1, 0}, {L"-v", L"1", L"2"}, switches));
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 1, 1, 0}, {L"-v", L"1", L"-v"}, switches));
			// Not enough non switch args
			Assert::AreEqual(false, cli::CheckMinArgs({0, 0, 1, 1, 0}, {L"-v", L"-v"}, switches));
			// Not enough args at all
			Assert::AreEqual(false, cli::CheckMinArgs({0, 0, 1, 1, 0}, {L"-v"}, switches));

			// min/max-0/1
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 0, 1, 0}, {L"-v", L"-v"}, switches));
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 0, 1, 0}, {L"-v", L"1", L"-v"}, switches));
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 0, 1, 0}, {L"-v"}, switches));

			// min/max-2/3
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 2, 3, 0}, {L"-v", L"1", L"2"}, switches));
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 2, 3, 0}, {L"-v", L"1", L"2", L"-v"}, switches));
			Assert::AreEqual(true, cli::CheckMinArgs({0, 0, 2, 3, 0}, {L"-v", L"1", L"2", L"3"}, switches));
			Assert::AreEqual(false, cli::CheckMinArgs({0, 0, 2, 3, 0}, {L"-v"}, switches));
			Assert::AreEqual(false, cli::CheckMinArgs({0, 0, 2, 3, 0}, {L"-v", L"1"}, switches));
			Assert::AreEqual(false, cli::CheckMinArgs({0, 0, 2, 3, 0}, {L"-v", L"1", L"-v"}, switches));
		}
	}; // namespace clitest
} // namespace clitest