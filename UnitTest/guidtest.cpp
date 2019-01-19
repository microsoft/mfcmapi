#include <../StdAfx.h>
#include <CppUnitTest.h>
#include <UnitTest/UnitTest.h>
#include <core/interpret/guid.h>

namespace guidtest
{
	TEST_CLASS(guidtest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_StringToGUID)
		{
			Assert::AreEqual(GUID_NULL, guid::StringToGUID(L""));
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"{00020D0A-0000-0000-C000-000000000046}"));
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"00020D0A-0000-0000-C000-000000000046"));
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"{00020D0A-0000-0000-C000-000000000046"));
			Assert::AreEqual(guid::IID_CAPONE_PROF, guid::StringToGUID(L"00020D0A-0000-0000-C000-000000000046}"));
			Assert::AreEqual(
				guid::IID_CAPONE_PROF, guid::StringToGUID(L"{0 0 020D0 A-0000-0000-C000-000000000046}  test"));
			Assert::AreEqual(guid::GUID_Dilkie, guid::StringToGUID(L"{53BC2EC0-D953-11CD-9752-00AA004AE40E}"));
			Assert::AreEqual(
				*LPGUID(pbGlobalProfileSectionGuid), guid::StringToGUID(L"{C8B0DB13-05AA-1A10-9BB0-00AA002FC45A}"));
			Assert::AreEqual(
				*LPGUID(pbGlobalProfileSectionGuid), guid::StringToGUID(L"13dbb0c8aa05101a9bb000aa002fc45a", true));
		}

		TEST_METHOD(Test_GUIDToString)
		{
			Assert::AreEqual(std::wstring(L"{00000000-0000-0000-0000-000000000000}"), guid::GUIDToString(nullptr));
			Assert::AreEqual(std::wstring(L"{00000000-0000-0000-0000-000000000000}"), guid::GUIDToString(GUID_NULL));
			Assert::AreEqual(
				std::wstring(L"{00020D0A-0000-0000-C000-000000000046}"), guid::GUIDToString(guid::IID_CAPONE_PROF));
			Assert::AreEqual(
				std::wstring(L"{00020D0A-0000-0000-C000-000000000046}"), guid::GUIDToString(&guid::IID_CAPONE_PROF));
			Assert::AreEqual(
				std::wstring(L"{53BC2EC0-D953-11CD-9752-00AA004AE40E}"), guid::GUIDToString(guid::GUID_Dilkie));
		}

		TEST_METHOD(Test_GUIDToStringAndName)
		{
			Assert::AreEqual(
				std::wstring(L"{00000000-0000-0000-0000-000000000000} = Unknown GUID"),
				guid::GUIDToStringAndName(nullptr));
			Assert::AreEqual(std::wstring(L"{00000000-0000-0000-0000-000000000000}"), guid::GUIDToString(GUID_NULL));
			Assert::AreEqual(
				std::wstring(L"{00020D0A-0000-0000-C000-000000000046} = IID_CAPONE_PROF"),
				guid::GUIDToStringAndName(guid::IID_CAPONE_PROF));
			Assert::AreEqual(
				std::wstring(L"{00020D0A-0000-0000-C000-000000000046} = IID_CAPONE_PROF"),
				guid::GUIDToStringAndName(&guid::IID_CAPONE_PROF));
			Assert::AreEqual(
				std::wstring(L"{53BC2EC0-D953-11CD-9752-00AA004AE40E} = GUID_Dilkie"),
				guid::GUIDToStringAndName(guid::GUID_Dilkie));
		}

		TEST_METHOD(Test_GUIDNameToGUID)
		{
			Assert::AreEqual(GUID_NULL, *guid::GUIDNameToGUID(L"GUID_NULL", false));
			Assert::AreEqual(
				guid::IID_CAPONE_PROF, *guid::GUIDNameToGUID(L"{00020D0A-0000-0000-C000-000000000046}", false));
			Assert::AreEqual(guid::IID_CAPONE_PROF, *guid::GUIDNameToGUID(L"IID_CAPONE_PROF", false));
		}
	};
} // namespace guidtest