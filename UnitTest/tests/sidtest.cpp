#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/interpret/sid.h>
#include <core/utility/strings.h>

namespace sidtest
{
	TEST_CLASS(sidtest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_LookupAccountSid)
		{
			auto nullAccount = sid::LookupAccountSid({});
			Assert::AreEqual(std::wstring{L"(no domain)"}, nullAccount.getDomain());
			Assert::AreEqual(std::wstring{L"(no name)"}, nullAccount.getName());

			auto invalidAccount = sid::LookupAccountSid({12});
			Assert::AreEqual(std::wstring{L"(no domain)"}, invalidAccount.getDomain());
			Assert::AreEqual(std::wstring{L"(no name)"}, invalidAccount.getName());

			auto simpleSidBin = strings::HexStringToBin(L"010500000000000515000000A065CF7E784B9B5FE77C8770091C0100");
			auto simpleSidAccount = sid::LookupAccountSid(simpleSidBin);
			Assert::AreEqual(std::wstring{L"(no domain)"}, simpleSidAccount.getDomain());
			Assert::AreEqual(std::wstring{L"(no name)"}, simpleSidAccount.getName());

			auto authenticatedUsersSidBin = strings::HexStringToBin(L"01 01 000000000005 0B000000");
			auto authenticatedUsersSidAccount = sid::LookupAccountSid(authenticatedUsersSidBin);
			Assert::AreEqual(std::wstring{L"NT AUTHORITY"}, authenticatedUsersSidAccount.getDomain());
			Assert::AreEqual(std::wstring{L"Authenticated Users"}, authenticatedUsersSidAccount.getName());
		}

		TEST_METHOD(Test_SidAccount)
		{
			auto emptyAccount = sid::SidAccount{};
			Assert::AreEqual(std::wstring{L"(no domain)"}, emptyAccount.getDomain());
			Assert::AreEqual(std::wstring{L"(no name)"}, emptyAccount.getName());

			auto account = sid::SidAccount{L"foo", L"bar"};
			Assert::AreEqual(std::wstring{L"foo"}, account.getDomain());
			Assert::AreEqual(std::wstring{L"bar"}, account.getName());

			auto account2 = sid::SidAccount{std::wstring{L"foo"}, std::wstring{L"bar"}};
			Assert::AreEqual(std::wstring{L"foo"}, account2.getDomain());
			Assert::AreEqual(std::wstring{L"bar"}, account2.getName());
		}
	};
} // namespace sidtest