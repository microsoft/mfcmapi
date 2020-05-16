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
				proptags::TagToString(PR_SUBJECT, nullptr, false, true));
			unittest::AreEqualEx(
				L"Tag: 0x0037001E\r\n"
				L"Type: PT_STRING8\r\n"
				L"Property Name: PR_SUBJECT\r\n"
				L"Other Names: PR_SUBJECT_A, ptagSubject, PR_SUBJECT_W, PidTagSubject\r\n"
				L"DASL: http://schemas.microsoft.com/mapi/proptag/0x0037001E",
				proptags::TagToString(PR_SUBJECT, nullptr, false, false));
			unittest::AreEqualEx(L"PR_SUBJECT", proptags::PropTagToPropName(PR_SUBJECT, false).bestGuess);
			unittest::AreEqualEx(L"PR_SUBJECT_A, ptagSubject, PR_SUBJECT_W, PidTagSubject", proptags::PropTagToPropName(PR_SUBJECT, false).otherMatches);
		}
	};
} // namespace proptagTest