#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>

namespace addin
{
	// imports for testing
	int _cdecl compareTypes(_In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareTags(_In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareNameID(_In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareSmartViewParser(_In_ const void* a1, _In_ const void* a2) noexcept;
} // namespace addin

namespace addintest
{
	int _cdecl testCompareTypes(_In_ const NAME_ARRAY_ENTRY& a1, _In_ const NAME_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareTypes(reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareTags(_In_ const NAME_ARRAY_ENTRY_V2& a1, _In_ const NAME_ARRAY_ENTRY_V2& a2) noexcept
	{
		return addin::compareTags(reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareNameID(_In_ const NAMEID_ARRAY_ENTRY& a1, _In_ const NAMEID_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareNameID(reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareSmartViewParser(
		_In_ const SMARTVIEW_PARSER_ARRAY_ENTRY& a1,
		_In_ const SMARTVIEW_PARSER_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareSmartViewParser(reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	TEST_CLASS(addinTest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_compareTypes)
		{
			Assert::AreEqual(0, testCompareTypes({1, L"one"}, {1, L"one"}));
			Assert::AreEqual(-1, testCompareTypes({1, L"one"}, {2, L"two"}));
			Assert::AreEqual(1, testCompareTypes({2, L"two"}, {1, L"one"}));

			Assert::AreEqual(-1, testCompareTypes({1, L"a"}, {1, L"b"}));
			Assert::AreEqual(1, testCompareTypes({1, L"b"}, {1, L"a"}));

			// Same name gets no special treatment
			Assert::AreEqual(-1, testCompareTypes({1, L"one"}, {2, L"one"}));
			Assert::AreEqual(1, testCompareTypes({2, L"one"}, {1, L"one"}));
		}

		TEST_METHOD(Test_compareTags)
		{
			Assert::AreEqual(0, testCompareTags({1, 0, L"one"}, {1, 0, L"one"}));
			Assert::AreEqual(-1, testCompareTags({1, 0, L"one"}, {2, 0, L"two"}));
			Assert::AreEqual(1, testCompareTags({2, 0, L"two"}, {1, 0, L"one"}));

			Assert::AreEqual(-1, testCompareTags({1, 0, L"a"}, {1, 0, L"b"}));
			Assert::AreEqual(-1, testCompareTags({1, 0, L"a"}, {1, 1, L"b"}));

			// Sort order field doesn't actually affect this particular sort
			Assert::AreEqual(-1, testCompareTags({1, 1, L"a"}, {1, 0, L"b"}));
			Assert::AreEqual(1, testCompareTags({1, 0, L"b"}, {1, 0, L"a"}));

			// Same name compares equal
			Assert::AreEqual(0, testCompareTags({1, 0, L"one"}, {2, 0, L"one"}));
		}

		TEST_METHOD(Test_compareNameID)
		{
			Assert::AreEqual(
				0,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"}));
			Assert::AreEqual(
				-1,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{2, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"}));
			Assert::AreEqual(
				1,
				testCompareNameID(
					{2, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"}));
			Assert::AreEqual(
				-1,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{1, &guid::PSETID_Note, L"a", PT_SYSTIME, L"aa"}));
			Assert::AreEqual(
				-1,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"},
					{1, &guid::PSETID_Meeting, L"b", PT_SYSTIME, L"bb"}));
			Assert::AreEqual(
				1,
				testCompareNameID(
					{1, &guid::PSETID_Meeting, L"b", PT_SYSTIME, L"bb"},
					{1, &guid::PSETID_Meeting, L"a", PT_SYSTIME, L"aa"}));
		}

		TEST_METHOD(Test_CompareSmartViewParser)
		{
			Assert::AreEqual(
				0, testCompareSmartViewParser({1, parserType::REPORTTAG, false}, {1, parserType::REPORTTAG, false}));
			Assert::AreEqual(
				-1, testCompareSmartViewParser({1, parserType::REPORTTAG, false}, {2, parserType::REPORTTAG, false}));
			Assert::AreEqual(
				1, testCompareSmartViewParser({2, parserType::REPORTTAG, false}, {1, parserType::REPORTTAG, false}));

			Assert::AreEqual(
				-1, testCompareSmartViewParser({1, parserType::ENTRYID, false}, {1, parserType::REPORTTAG, false}));
			Assert::AreEqual(
				1, testCompareSmartViewParser({1, parserType::REPORTTAG, false}, {1, parserType::ENTRYID, false}));

			Assert::AreEqual(
				-1, testCompareSmartViewParser({1, parserType::REPORTTAG, false}, {1, parserType::REPORTTAG, true}));
			Assert::AreEqual(
				1, testCompareSmartViewParser({1, parserType::REPORTTAG, true}, {1, parserType::REPORTTAG, false}));
		}
	};
} // namespace addintest