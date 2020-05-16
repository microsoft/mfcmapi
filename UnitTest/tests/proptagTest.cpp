#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/interpret/proptags.h>

namespace proptagTest
{
	TEST_CLASS(proptagTest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_TagToString)
		{
			unittest::AreEqualEx(
				L"0x0037001E (PT_STRING8): PR_SUBJECT: (PR_SUBJECT_A, ptagSubject, PR_SUBJECT_W, PidTagSubject)",
				proptags::TagToString(PR_SUBJECT_A, nullptr, false, true));
			unittest::AreEqualEx(
				L"Tag: 0x0037001E\r\n"
				L"Type: PT_STRING8\r\n"
				L"Property Name: PR_SUBJECT\r\n"
				L"Other Names: PR_SUBJECT_A, ptagSubject, PR_SUBJECT_W, PidTagSubject\r\n"
				L"DASL: http://schemas.microsoft.com/mapi/proptag/0x0037001E",
				proptags::TagToString(PR_SUBJECT_A, nullptr, false, false));
		}

		TEST_METHOD(Test_PropTagToPropName)
		{
			unittest::AreEqualEx(L"PR_SUBJECT", proptags::PropTagToPropName(PR_SUBJECT_A, false).bestGuess);
			unittest::AreEqualEx(
				L"PR_SUBJECT_A, ptagSubject, PR_SUBJECT_W, PidTagSubject",
				proptags::PropTagToPropName(PR_SUBJECT_A, false).otherMatches);

			unittest::AreEqualEx(L"PR_RW_RULES_STREAM", proptags::PropTagToPropName(0x68020102, false).bestGuess);
			unittest::AreEqualEx(
				L"PidTagRwRulesStream, PR_OAB_CONTAINER_GUID, PR_OAB_CONTAINER_GUID_W, "
				L"PidTagOfflineAddressBookContainerGuid, PidTagSenderTelephoneNumber",
				proptags::PropTagToPropName(0x68020102, false).otherMatches);
			unittest::AreEqualEx(L"PR_RW_RULES_STREAM", proptags::PropTagToPropName(0x6802000A, false).bestGuess);
		}
	};
} // namespace proptagTest