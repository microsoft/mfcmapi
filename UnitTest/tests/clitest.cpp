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
	cmdmodeHelpFull,
};

namespace Microsoft
{
	namespace VisualStudio
	{
		namespace CppUnitTestFramework
		{
			//template <>
			//inline std::wstring ToString<cli::OptParser<switchEnum, modeEnum, flagsEnum>>(
			//	const cli::OptParser<switchEnum, modeEnum, flagsEnum>& q)
			//{
			//	return strings::format(L"%d-%d-%d-%d-%d", q.Switch, q.Mode, q.MinArgs, q.MaxArgs, q.ulOpt);
			//}

			void AreEqual(
				const cli::OptParser<switchEnum, modeEnum, flagsEnum>& expected,
				const cli::OptParser<switchEnum, modeEnum, flagsEnum>& actual,
				const wchar_t* message = nullptr,
				const __LineInfo* pLineInfo = nullptr)
			{
				auto eq = true;
				if (expected.Switch != actual.Switch || expected.Mode != actual.Mode ||
					expected.MinArgs != actual.MinArgs || expected.MaxArgs != actual.MaxArgs ||
					expected.ulOpt != actual.ulOpt)
				{
					eq = false;
				}

				if (!eq)
				{
					Logger::WriteMessage(strings::format(L"Switch: %d:%d\n", expected.Switch, actual.Switch).c_str());
					Logger::WriteMessage(strings::format(L"Mode: %d:%d\n", expected.Mode, actual.Mode).c_str());
					Logger::WriteMessage(
						strings::format(L"MinArgs: %d:%d\n", expected.MinArgs, actual.MinArgs).c_str());
					Logger::WriteMessage(
						strings::format(L"MaxArgs: %d:%d\n", expected.MaxArgs, actual.MaxArgs).c_str());
					Logger::WriteMessage(strings::format(L"ulOpt: %d:%d\n", expected.ulOpt, actual.ulOpt).c_str());
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
			Assert::AreEqual(std::vector<std::wstring>{}, cli::GetCommandLine(noarg.size(), noarg.data()));
			auto argv = std::vector<LPCSTR>{"app.exe", "-arg1", "-arg2"};
			Assert::AreEqual(
				std::vector<std::wstring>{L"-arg1", L"-arg2"}, cli::GetCommandLine(argv.size(), argv.data()));
		}

		TEST_METHOD(Test_ParseArgument)
		{
			const std::vector<cli::COMMANDLINE_SWITCH<switchEnum>> switches = {
				{switchNoSwitch, L""},
				{switchUnknown, L""},
				{switchHelp, L"?"},
				{switchVerbose, L"Verbose"},
				{switchSearch, L"Search"},
			};

			const std::vector<cli::OptParser<switchEnum, modeEnum, flagsEnum>> parsers = {
				{switchHelp, cmdmodeHelpFull, 0, 0, OPT_INITMFC},
				{switchVerbose, cmdmodeUnknown, 0, 0, OPT_VERBOSE | OPT_INITMFC},
				{switchSearch, cmdmodeUnknown, 0, 0, OPT_DOPARTIALSEARCH},
				{switchNoSwitch, cmdmodeUnknown, 0, 0, OPT_NOOPT},
			};

			Assert::AreEqual(int(switchHelp), int(cli::ParseArgument<switchEnum>(std::wstring{L"-?"}, switches)));
			Assert::AreEqual(int(switchVerbose), int(cli::ParseArgument<switchEnum>(std::wstring{L"-v"}, switches)));
			Assert::AreEqual(int(switchVerbose), int(cli::ParseArgument<switchEnum>(std::wstring{L"/v"}, switches)));
			Assert::AreEqual(int(switchVerbose), int(cli::ParseArgument<switchEnum>(std::wstring{L"\\v"}, switches)));
			Assert::AreEqual(
				int(switchVerbose), int(cli::ParseArgument<switchEnum>(std::wstring{L"-verbose"}, switches)));
			Assert::AreEqual(
				int(switchNoSwitch), int(cli::ParseArgument<switchEnum>(std::wstring{L"-verbosey"}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(cli::ParseArgument<switchEnum>(std::wstring{L"-va"}, switches)));
			Assert::AreEqual(
				int(switchNoSwitch), int(cli::ParseArgument<switchEnum>(std::wstring{L"-test"}, switches)));
			Assert::AreEqual(int(switchSearch), int(cli::ParseArgument<switchEnum>(std::wstring{L"/s"}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(cli::ParseArgument<switchEnum>(std::wstring{L""}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(cli::ParseArgument<switchEnum>(std::wstring{L"+v"}, switches)));
			Assert::AreEqual(int(switchNoSwitch), int(cli::ParseArgument<switchEnum>(std::wstring{L"-"}, switches)));

			AreEqual(parsers[0], cli::GetParser<switchEnum, modeEnum, flagsEnum>(switchHelp, parsers));
			AreEqual(parsers[1], cli::GetParser<switchEnum, modeEnum, flagsEnum>(switchVerbose, parsers));
			AreEqual({}, cli::GetParser<switchEnum, modeEnum, flagsEnum>(switchNoSwitch, parsers));
		}
	};
} // namespace clitest