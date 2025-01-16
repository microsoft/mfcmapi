#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/interpret/sid.h>
#include <core/utility/strings.h>

namespace sid
{
	std::wstring ACEToString(const std::vector<BYTE>& buf, aceType acetype);
}

namespace sidtest
{
	TEST_CLASS(sidtest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_GetTextualSid)
		{
			Assert::AreEqual(std::wstring{}, sid::GetTextualSid({}));
			auto nullAccount = sid::LookupAccountSid({});
			Assert::AreEqual(std::wstring{L"(no domain)"}, nullAccount.getDomain());
			Assert::AreEqual(std::wstring{L"(no name)"}, nullAccount.getName());

			Assert::AreEqual(std::wstring{}, sid::GetTextualSid({12}));
			auto invalidAccount = sid::LookupAccountSid({12});
			Assert::AreEqual(std::wstring{L"(no domain)"}, invalidAccount.getDomain());
			Assert::AreEqual(std::wstring{L"(no name)"}, invalidAccount.getName());

			auto simpleSidBin = strings::HexStringToBin(L"010500000000000515000000A065CF7E784B9B5FE77C8770091C0100");
			Assert::AreEqual(
				std::wstring{L"S-1-5-21-2127521184-1604012920-1887927527-72713"}, sid::GetTextualSid(simpleSidBin));
			auto simpleSidAccount = sid::LookupAccountSid(simpleSidBin);
			Assert::AreEqual(std::wstring{L"(no domain)"}, simpleSidAccount.getDomain());
			Assert::AreEqual(std::wstring{L"(no name)"}, simpleSidAccount.getName());

			Assert::AreEqual(
				std::wstring{L"S-1-000102030405-21-2127521184-1604012920-1887927527-72713"},
				sid::GetTextualSid(
					strings::HexStringToBin(L"010500010203040515000000A065CF7E784B9B5FE77C8770091C0100")));

			auto authenticatedUsersSidBin = strings::HexStringToBin(L"01 01 000000000005 0B000000");
			Assert::AreEqual(std::wstring{L"S-1-5-11"}, sid::GetTextualSid(authenticatedUsersSidBin));
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

		TEST_METHOD(Test_ACEToString)
		{
			// test ACEToString with a zero length vector
			unittest::AreEqualEx(std::wstring{L""}, ACEToString({}, sid::aceType::Container));
		}
	};
} // namespace sidtest