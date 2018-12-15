#include <StdAfx.h>
#include <CppUnitTest.h>
#include <Interpret/Sid.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

//namespace Microsoft
//{
//	namespace VisualStudio
//	{
//		namespace CppUnitTestFramework
//		{
//			template <> inline std::wstring ToString<GUID>(const GUID& t) { return guid::GUIDToString(t); }
//		} // namespace CppUnitTestFramework
//	} // namespace VisualStudio
//} // namespace Microsoft

namespace sidtest
{
	TEST_CLASS(sidtest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(Initialize)
		{
			// Set up our property arrays or nothing works
			addin::MergeAddInArrays();

			registry::RegKeys[registry::regkeyDO_SMART_VIEW].ulCurDWORD = 1;
			registry::RegKeys[registry::regkeyUSE_GETPROPLIST].ulCurDWORD = 1;
			registry::RegKeys[registry::regkeyPARSED_NAMED_PROPS].ulCurDWORD = 1;
			registry::RegKeys[registry::regkeyCACHE_NAME_DPROPS].ulCurDWORD = 1;

			strings::setTestInstance(GetModuleHandleW(L"UnitTest.dll"));
		}

		TEST_METHOD(Test_GetTextualSid)
		{
			Assert::AreEqual(std::wstring{}, sid::GetTextualSid(nullptr));
			auto simpleSidBin = strings::HexStringToBin(L"010500000000000515000000A065CF7E784B9B5FE77C8770091C0100");
			auto simpleSid = static_cast<PSID>(simpleSidBin.data());
			Assert::AreEqual(
				std::wstring{L"S-1-5-21-2127521184-1604012920-1887927527-72713"}, sid::GetTextualSid(simpleSid));
			auto simpleSidAccount = sid::LookupAccountSid(simpleSid);
			Assert::AreEqual(std::wstring{L"(no domain)"}, simpleSidAccount.getDomain());
			Assert::AreEqual(std::wstring{L"(no name)"}, simpleSidAccount.getName());

			auto longerSidBin = strings::HexStringToBin(L"010500010203040515000000A065CF7E784B9B5FE77C8770091C0100");
			auto longerSid = static_cast<PSID>(longerSidBin.data());
			Assert::AreEqual(
				std::wstring{L"S-1-000102030405-21-2127521184-1604012920-1887927527-72713"},
				sid::GetTextualSid(longerSid));

			auto authenticatedUsersSidBin = strings::HexStringToBin(L"01 01 000000000005 0B000000");
			auto authenticatedUsersSid = static_cast<PSID>(authenticatedUsersSidBin.data());
			Assert::AreEqual(std::wstring{L"S-1-5-11"}, sid::GetTextualSid(authenticatedUsersSid));
			auto authenticatedUsersSidAccount = sid::LookupAccountSid(authenticatedUsersSid);
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
		}
	};
} // namespace sidtest