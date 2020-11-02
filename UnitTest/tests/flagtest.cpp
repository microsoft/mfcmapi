#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace flagtest
{
	TEST_CLASS(flagtest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_InterpretFlags)
		{
			// flagVALUE
			unittest::AreEqualEx(L"BMR_EQZ", flags::InterpretFlags(flagBitmask, BMR_EQZ));
			unittest::AreEqualEx(L"BMR_NEZ", flags::InterpretFlags(flagBitmask, BMR_NEZ));

			// flagFLAG with null flagVALUE
			unittest::AreEqualEx(
				L"MSG_DEFAULTS_NONE", flags::InterpretFlags(PROP_ID(PR_CERT_DEFAULTS), MSG_DEFAULTS_NONE));
			unittest::AreEqualEx(
				L"MSG_DEFAULTS_FOR_FORMAT | MSG_DEFAULTS_SEND_CERT",
				flags::InterpretFlags(PROP_ID(PR_CERT_DEFAULTS), MSG_DEFAULTS_FOR_FORMAT | MSG_DEFAULTS_SEND_CERT));
			unittest::AreEqualEx(
				L"MSG_DEFAULTS_GLOBAL | 0x8",
				flags::InterpretFlags(PROP_ID(PR_CERT_DEFAULTS), MSG_DEFAULTS_GLOBAL | 0x8));

			// flagFLAG mixed with flagVALUE
			unittest::AreEqualEx(
				L"MAPI_P1 | MAPI_SUBMITTED | MAPI_ORIG",
				flags::InterpretFlags(PROP_ID(PR_RECIPIENT_TYPE), MAPI_P1 | MAPI_SUBMITTED | MAPI_ORIG));
			unittest::AreEqualEx(
				L"MAPI_P1 | MAPI_TO", flags::InterpretFlags(PROP_ID(PR_RECIPIENT_TYPE), MAPI_P1 | MAPI_TO));
			unittest::AreEqualEx(
				L"MAPI_SUBMITTED | 0x5", flags::InterpretFlags(PROP_ID(PR_RECIPIENT_TYPE), MAPI_SUBMITTED | 0x5));
			unittest::AreEqualEx(
				L"MAPI_SUBMITTED | 0x20000000 | 0x5", flags::InterpretFlags(PROP_ID(PR_RECIPIENT_TYPE), MAPI_SUBMITTED | 0x20000005));

			// flagFLAG with no null
			unittest::AreEqualEx(L"0x0", flags::InterpretFlags(PROP_ID(PR_SUBMIT_FLAGS), 0));
			unittest::AreEqualEx(
				L"SUBMITFLAG_LOCKED", flags::InterpretFlags(PROP_ID(PR_SUBMIT_FLAGS), SUBMITFLAG_LOCKED));

			// FLAG_ENTRY3RDBYTE and FLAG_ENTRY4THBYTE
			unittest::AreEqualEx(
				L"DTE_FLAG_REMOTE_VALID | Remote: DT_MAILUSER | Local: DT_SEC_DISTLIST",
				flags::InterpretFlags(
					PROP_ID(PR_DISPLAY_TYPE_EX),
					DTE_FLAG_REMOTE_VALID | DTE_REMOTE(DT_ROOM) | DTE_LOCAL(DT_SEC_DISTLIST)));

			// FLAG_ENTRYHIGHBYTES
			unittest::AreEqualEx(
				L"CPU: MAPIFORM_CPU_X86 | MAPIFORM_OS_WIN_95",
				flags::InterpretFlags(
					PROP_ID(PR_FORM_HOST_MAP), MAPIFORM_PLATFORM(MAPIFORM_CPU_X86, MAPIFORM_OS_WIN_95)));

			// NON_PROP_FLAG_ENTRYLOWERNIBBLE
			unittest::AreEqualEx(
				L"BFLAGS_MASK_OUTLOOK | Type: BFLAGS_INTERNAL_MAILUSER | Home Fax",
				flags::InterpretFlags(flagWABEntryIDType, BFLAGS_MASK_OUTLOOK | BFLAGS_INTERNAL_MAILUSER | 0x10));
		}

		TEST_METHOD(Test_AllFlagsToString)
		{
			unittest::AreEqualEx(L"", flags::AllFlagsToString(NULL, true));
			unittest::AreEqualEx(L"", flags::AllFlagsToString(0xFFFFFFFF, true));
			unittest::AreEqualEx(
				L"\r\n0x10000000 MAPI_P1\r\n"
				L"0x80000000 MAPI_SUBMITTED\r\n"
				L"0x00000000 MAPI_ORIG\r\n"
				L"0x00000001 MAPI_TO\r\n"
				L"0x00000002 MAPI_CC\r\n"
				L"0x00000003 MAPI_BCC",
				flags::AllFlagsToString(PROP_ID(PR_RECIPIENT_TYPE), true));
			unittest::AreEqualEx(
				L"\r\n268435456 MAPI_P1\r\n"
				L"-2147483648 MAPI_SUBMITTED\r\n"
				L"    0 MAPI_ORIG\r\n"
				L"    1 MAPI_TO\r\n"
				L"    2 MAPI_CC\r\n"
				L"    3 MAPI_BCC",
				flags::AllFlagsToString(PROP_ID(PR_RECIPIENT_TYPE), false));
		}
	};
} // namespace flagtest