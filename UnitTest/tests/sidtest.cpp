#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/interpret/sid.h>
#include <core/utility/strings.h>

namespace sid
{
	std::wstring ACEToString(_In_opt_ void* pACE, aceType acetype);
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
			unittest::AreEqualEx(std::wstring{L""}, ACEToString(nullptr, sid::aceType::Container));

			auto aceAllowBin =
				strings::HexStringToBin(L"00092400a9081200010500000000000515000000371a6c07352f372aad20fa5b01930100");
			unittest::AreEqualEx(
				std::wstring{
					L"Account: (no domain)\\(no name)\r\n"
					L"SID: S-1-5-21-124525111-708259637-1543119021-103169\r\n"
					L"Access Type: 0x00000000 = ACCESS_ALLOWED_ACE_TYPE\r\n"
					L"Access Flags: 0x00000009 = OBJECT_INHERIT_ACE | INHERIT_ONLY_ACE\r\n"
					L"Access Mask: 0x001208A9 = fsdrightListContents | fsdrightReadProperty | fsdrightExecute | "
					L"fsdrightReadAttributes | fsdrightViewItem | fsdrightReadControl | fsdrightSynchronize"},
				ACEToString(aceAllowBin.data(), sid::aceType::Container));

			auto aceDenyBin = strings::HexStringToBin(L"01 09 1400 a9081200 01 01 000000000005 0B000000");
			unittest::AreEqualEx(
				std::wstring{L"Account: NT AUTHORITY\\Authenticated Users\r\n"
							 L"SID: S-1-5-11\r\n"
							 L"Access Type: 0x00000001 = ACCESS_DENIED_ACE_TYPE\r\n"
							 L"Access Flags: 0x00000009 = OBJECT_INHERIT_ACE | INHERIT_ONLY_ACE\r\n"
							 L"Access Mask: 0x001208A9 = "},
				ACEToString(aceDenyBin.data(), sid::aceType{3}));

			auto aceAllowObjectBin = strings::HexStringToBin(L"05 1f 3800 a9081200 ffffffff"
															 L"0A0D0200-0000-0000-C000-000000000046"
															 L"C02EBC53-53D9-CD11-9752-00AA004AE40E"
															 L"FF 01 000000000005 0B000000");
			unittest::AreEqualEx(
				std::wstring{
					L"Account: (no domain)\\(no name)\r\n"
					L"SID: (no SID)\r\n"
					L"Access Type: 0x00000005 = ACCESS_ALLOWED_OBJECT_ACE_TYPE\r\n"
					L"Access Flags: 0x0000001F = OBJECT_INHERIT_ACE | CONTAINER_INHERIT_ACE | NO_PROPAGATE_INHERIT_ACE "
					L"| INHERIT_ONLY_ACE | INHERITED_ACE\r\n"
					L"Access Mask: 0x001208A9 = fsdrightReadBody | fsdrightReadProperty | fsdrightExecute | "
					L"fsdrightReadAttributes | fsdrightViewItem | fsdrightReadControl | fsdrightSynchronize\r\n"
					L"ObjectType: \r\n"
					L"{00020D0A-0000-0000-C000-000000000046} = IID_CAPONE_PROF\r\n"
					L"InheritedObjectType: \r\n"
					L"{53BC2EC0-D953-11CD-9752-00AA004AE40E} = GUID_Dilkie\r\n"
					L"Flags: 0xFFFFFFFF"},
				ACEToString(aceAllowObjectBin.data(), sid::aceType::Message));

			auto aceDenyObjectBin = strings::HexStringToBin(L"06 1f 3800 03000000 ffffffff"
															L"0A0D0200-0000-0000-C000-000000000046"
															L"C02EBC53-53D9-CD11-9752-00AA004AE40E"
															L"01 01 000000000005 0B000000");
			unittest::AreEqualEx(
				std::wstring{
					L"Account: NT AUTHORITY\\Authenticated Users\r\n"
					L"SID: S-1-5-11\r\n"
					L"Access Type: 0x00000006 = ACCESS_DENIED_OBJECT_ACE_TYPE\r\n"
					L"Access Flags: 0x0000001F = OBJECT_INHERIT_ACE | CONTAINER_INHERIT_ACE | NO_PROPAGATE_INHERIT_ACE "
					L"| INHERIT_ONLY_ACE | INHERITED_ACE\r\n"
					L"Access Mask: 0x00000003 = fsdrightFreeBusySimple | fsdrightFreeBusyDetailed\r\n"
					L"ObjectType: \r\n"
					L"{00020D0A-0000-0000-C000-000000000046} = IID_CAPONE_PROF\r\n"
					L"InheritedObjectType: \r\n"
					L"{53BC2EC0-D953-11CD-9752-00AA004AE40E} = GUID_Dilkie\r\n"
					L"Flags: 0xFFFFFFFF"},
				ACEToString(aceDenyObjectBin.data(), sid::aceType::FreeBusy));
		}

		TEST_METHOD(Test_SDToString)
		{
			const auto nullsd = SDToString({}, sid::aceType::Container);
			Assert::AreEqual(std::wstring{L"This is not a valid security descriptor."}, nullsd.dacl);
			Assert::AreEqual(std::wstring{L""}, nullsd.info);

			const auto invalid =
				SDToString(strings::HexStringToBin(L"B606B07ABB6079AB2082C760"), sid::aceType::Container);
			Assert::AreEqual(std::wstring{L"This is not a valid security descriptor."}, invalid.dacl);
			Assert::AreEqual(std::wstring{L""}, invalid.info);

			const auto sd = SDToString(
				strings::HexStringToBin(L"0800030000000000010007801C000000280000000000000014000000020008000000000001010"
										L"000000000051200000001020000000000052000000020020000"),
				sid::aceType::Container);
			Assert::AreEqual(std::wstring{L""}, sd.dacl);
			Assert::AreEqual(std::wstring{L"0x0"}, sd.info);

			const auto sd1 = SDToString(
				strings::HexStringToBin(
					L"08000300000000000100078064000000700000000000000014000000020050000200000001092400BF0F1F00010500000"
					L"000000515000000271A6C07352F372AAD20FA5BAA830B0001022400BFC91F00010500000000000515000000271A6C0735"
					L"2F372AAD20FA5BAA830B0001010000000000051200000001020000000000052000000020020000"),
				sid::aceType::Container);
			unittest::AreEqualEx(
				std::wstring{
					L"Account: (no domain)\\(no name)\r\n"
					L"SID: S-1-5-21-124525095-708259637-1543119021-754602\r\n"
					L"Access Type: 0x00000001 = ACCESS_DENIED_ACE_TYPE\r\n"
					L"Access Flags: 0x00000009 = OBJECT_INHERIT_ACE | INHERIT_ONLY_ACE\r\n"
					L"Access Mask: 0x001F0FBF = fsdrightListContents | fsdrightCreateItem | fsdrightCreateContainer | "
					L"fsdrightReadProperty | fsdrightWriteProperty | fsdrightExecute | fsdrightReadAttributes | "
					L"fsdrightWriteAttributes | fsdrightViewItem | fsdrightWriteSD | fsdrightDelete | "
					L"fsdrightWriteOwner | fsdrightReadControl | fsdrightSynchronize | 0x600\r\n"
					L"Account: (no domain)\\(no name)\r\n"
					L"SID: S-1-5-21-124525095-708259637-1543119021-754602\r\n"
					L"Access Type: 0x00000001 = ACCESS_DENIED_ACE_TYPE\r\n"
					L"Access Flags: 0x00000002 = CONTAINER_INHERIT_ACE\r\n"
					L"Access Mask: 0x001FC9BF = fsdrightListContents | fsdrightCreateItem | fsdrightCreateContainer | "
					L"fsdrightReadProperty | fsdrightWriteProperty | fsdrightExecute | fsdrightReadAttributes | "
					L"fsdrightWriteAttributes | fsdrightViewItem | fsdrightOwner | fsdrightContact | fsdrightWriteSD | "
					L"fsdrightDelete | fsdrightWriteOwner | fsdrightReadControl | fsdrightSynchronize"},
				sd1.dacl);
			unittest::AreEqualEx(std::wstring{L"0x0"}, sd1.info);
		}
	};
} // namespace sidtest