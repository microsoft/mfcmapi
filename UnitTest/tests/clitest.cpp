#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/cli.h>

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
			enum switchEnum
			{
				switchNoSwitch = 0, // not a switch
				switchUnknown, // unknown switch
				switchHelp, // '-h'
				switchVerbose, // '-v'
				switchSearch, // '-s'
			};
			const std::vector<cli::COMMANDLINE_SWITCH<switchEnum>> switches = {
				{switchNoSwitch, L""},
				{switchUnknown, L""},
				{switchHelp, L"?"},
				{switchVerbose, L"Verbose"},
				{switchSearch, L"Search"},
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
		}
	};
} // namespace clitest