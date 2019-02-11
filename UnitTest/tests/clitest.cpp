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
			Assert::AreEqual(int(cli::switchEnum::switchHelp), int(ParseArgument(std::wstring{L"-?"}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchVerbose), int(ParseArgument(std::wstring{L"-v"}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchVerbose), int(ParseArgument(std::wstring{L"/v"}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchVerbose), int(ParseArgument(std::wstring{L"\\v"}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchVerbose), int(ParseArgument(std::wstring{L"-verbose"}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchNoSwitch), int(ParseArgument(std::wstring{L"-verbosey"}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchNoSwitch), int(ParseArgument(std::wstring{L"-va"}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchNoSwitch), int(ParseArgument(std::wstring{L"-test"}, cli::parsers)));
			Assert::AreEqual(int(cli::switchEnum::switchNoSwitch), int(ParseArgument(std::wstring{L""}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchNoSwitch), int(ParseArgument(std::wstring{L"+v"}, cli::parsers)));
			Assert::AreEqual(
				int(cli::switchEnum::switchNoSwitch), int(ParseArgument(std::wstring{L"-"}, cli::parsers)));
		}

		TEST_METHOD(Test_GetParser)
		{
			AreEqual(cli::helpParser, GetParser(cli::switchEnum::switchHelp, cli::parsers));
			AreEqual(cli::verboseParser, GetParser(cli::switchEnum::switchVerbose, cli::parsers));
			AreEqual(cli::noSwitchParser, GetParser(cli::switchEnum::switchNoSwitch, cli::parsers));
			AreEqual({}, GetParser(cli::switchEnum::switchNoSwitch, cli::parsers));
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

		TEST_METHOD(Test_CheckMinArgs)
		{
			// min/max-0/0
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 0, 0, 0}, {L"-v"}, cli::parsers));
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 0, 0, 0}, {L"-v", L"-v"}, cli::parsers));
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 0, 0, 0}, {L"-v", L"1"}, cli::parsers));

			// min/max-1/1
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 1, 1, 0}, {L"-v", L"1"}, cli::parsers));
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 1, 1, 0}, {L"-v", L"1", L"2"}, cli::parsers));
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 1, 1, 0}, {L"-v", L"1", L"-v"}, cli::parsers));
			// Not enough non switch args
			Assert::AreEqual(false, cli::CheckMinArgs({0, L"", 0, 1, 1, 0}, {L"-v", L"-v"}, cli::parsers));
			// Not enough args at all
			Assert::AreEqual(false, cli::CheckMinArgs({0, L"", 0, 1, 1, 0}, {L"-v"}, cli::parsers));

			// min/max-0/1
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 0, 1, 0}, {L"-v", L"-v"}, cli::parsers));
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 0, 1, 0}, {L"-v", L"1", L"-v"}, cli::parsers));
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 0, 1, 0}, {L"-v"}, cli::parsers));

			// min/max-2/3
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 2, 3, 0}, {L"-v", L"1", L"2"}, cli::parsers));
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 2, 3, 0}, {L"-v", L"1", L"2", L"-v"}, cli::parsers));
			Assert::AreEqual(true, cli::CheckMinArgs({0, L"", 0, 2, 3, 0}, {L"-v", L"1", L"2", L"3"}, cli::parsers));
			Assert::AreEqual(false, cli::CheckMinArgs({0, L"", 0, 2, 3, 0}, {L"-v"}, cli::parsers));
			Assert::AreEqual(false, cli::CheckMinArgs({0, L"", 0, 2, 3, 0}, {L"-v", L"1"}, cli::parsers));
			Assert::AreEqual(false, cli::CheckMinArgs({0, L"", 0, 2, 3, 0}, {L"-v", L"1", L"-v"}, cli::parsers));
		}
	}; // namespace clitest
} // namespace clitest