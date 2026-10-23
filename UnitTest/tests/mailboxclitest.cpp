#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <MrMapi/MMcli.h>

namespace mailboxclitest
{
	TEST_CLASS(mailboxclitest){public : TEST_CLASS_INITIALIZE(initialize){unittest::init();
}
TEST_METHOD_INITIALIZE(clearOptions) { ClearOptions(); }
TEST_METHOD_CLEANUP(restoreOptions) { ClearOptions(); }

TEST_METHOD(AcceptsStoreAndFolderInspectionModes)
{
	const std::vector<std::vector<std::wstring>> actions = {
		{L"-Contents"},
		{L"-HiddenContents"},
		{L"-Acl"},
		{L"-Rules"},
		{L"-ChildFolders"},
		{L"-FID"},
		{L"-Store"},
		{L"-Folder", L"13"},
		{L"-Size"},
		{L"-ReceiveFolder"},
		{L"-SearchState"},
		{L"-NamedProps"}};
	for (const auto& action : actions)
	{
		ClearOptions();
		auto args = BaseArguments();
		args.insert(args.end(), action.begin(), action.end());
		const auto options = Parse(args);
		Assert::IsTrue(cli::ValidateMailboxOptions(options).empty());
		Assert::IsTrue((options.flags & cli::OPT_NEEDSTORE) != 0);
		Assert::IsTrue((options.flags & cli::OPT_NEEDMAPILOGON) != 0);
	}
}

TEST_METHOD(ContentsUsesDefaultInboxAndAllowsExplicitPath)
{
	auto args = BaseArguments();
	args.push_back(L"-Contents");
	Assert::IsTrue(cli::ValidateMailboxOptions(Parse(args)).empty());
	Assert::AreEqual(std::wstring(L"13"), cli::switchFolder[0]);
	ClearOptions();
	args.insert(args.end(), {L"-Folder", L"Inbox\\Child"});
	Assert::IsTrue(cli::ValidateMailboxOptions(Parse(args)).empty());
	Assert::AreEqual(std::wstring(L"Inbox\\Child"), cli::switchFolder[0]);
}

TEST_METHOD(RequiresExplicitNonemptyProfile)
{
	auto options = Parse({L"-Mailbox", L"user@example.test", L"-Contents"});
	Assert::IsFalse(cli::ValidateMailboxOptions(options).empty());
	ClearOptions();
	options = Parse({L"-Mailbox", L"user@example.test", L"-Profile", L"", L"-Contents"});
	Assert::IsFalse(cli::ValidateMailboxOptions(options).empty());
}

TEST_METHOD(RejectsInvalidAddressAndUnsupportedAction)
{
	auto options = Parse({L"-Mailbox", L"not-an-address", L"-Profile", L"Service", L"-Contents"});
	Assert::IsFalse(cli::ValidateMailboxOptions(options).empty());
	ClearOptions();
	auto args = BaseArguments();
	args.push_back(L"-Accounts");
	options = Parse(args);
	Assert::IsFalse(cli::ValidateMailboxOptions(options).empty());
}

TEST_METHOD(RejectsEveryStoreRetargetingSyntax)
{
	const std::vector<std::vector<std::wstring>> actions = {
		{L"-Store", L"0"}, {L"-ReceiveFolder", L"0"}, {L"-Folder", L"#0\\Inbox"}, {L"-Folder", L"#AABB\\Inbox"}};
	for (const auto& action : actions)
	{
		ClearOptions();
		auto args = BaseArguments();
		args.insert(args.end(), action.begin(), action.end());
		Assert::IsFalse(cli::ValidateMailboxOptions(Parse(args)).empty());
	}
}

TEST_METHOD(CommandsWithoutMailboxKeepExistingValidation)
{
	const auto options = Parse({L"-Store", L"0"});
	Assert::IsTrue(cli::ValidateMailboxOptions(options).empty());
	Assert::IsFalse(cli::switchMailbox.isSet());
	Assert::AreEqual(static_cast<int>(cli::cmdmodeStoreProperties), options.mode);
	Assert::AreEqual(std::wstring(L"0"), cli::switchStore[0]);
}

private:
void ClearOptions()
{
	for (auto option : cli::g_options)
		option->clear();
	cli::switchUnswitched.clear();
}

std::vector<std::wstring> BaseArguments() { return {L"-Mailbox", L"user@example.test", L"-Profile", L"Service"}; }

cli::OPTIONS Parse(const std::vector<std::wstring>& arguments)
{
	auto args = std::deque<std::wstring>{arguments.begin(), arguments.end()};
	cli::OPTIONS options{};
	cli::ParseArgs(options, args, cli::g_options);
	cli::PostParseCheck(options);
	return options;
}
}
;
} // namespace mailboxclitest
