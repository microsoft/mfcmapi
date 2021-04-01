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

	void SortFlagArray(_In_count_(ulFlags) LPFLAG_ARRAY_ENTRY lpFlags, _In_ size_t ulFlags) noexcept;
	void AppendFlagIfNotDupe(std::vector<FLAG_ARRAY_ENTRY>& target, FLAG_ARRAY_ENTRY source);
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

		void testFA(std::wstring testName, const FLAG_ARRAY_ENTRY& fa1, const FLAG_ARRAY_ENTRY& fa2)
		{
			Assert::AreEqual(fa1.ulFlagName, fa2.ulFlagName, (testName + L"-ulFlagName").c_str());
			Assert::AreEqual(fa1.lFlagValue, fa2.lFlagValue, (testName + L"-lFlagValue").c_str());
			Assert::AreEqual(fa1.ulFlagType, fa2.ulFlagType, (testName + L"-ulFlagType").c_str());
			Assert::AreEqual(fa1.lpszName, fa2.lpszName, (testName + L"-lpszName").c_str());
		}

		TEST_METHOD(Test_SortFlagArray)
		{
			FLAG_ARRAY_ENTRY flagArray[] = {
				{2, 1, flagVALUE, L"b one"},
				{1, 2, flagVALUE, L"two"},
				{2, 2, flagVALUE, L"a two"},
				{1, 1, flagVALUE, L"one"},
			};

			// Do a stable sort on ulFlagName (first member)
			addin::SortFlagArray(flagArray, _countof(flagArray));

			testFA(L"0", flagArray[0], {1, 2, flagVALUE, L"two"});
			testFA(L"1", flagArray[1], {1, 1, flagVALUE, L"one"});
			testFA(L"2", flagArray[2], {2, 1, flagVALUE, L"b one"});
			testFA(L"3", flagArray[3], {2, 2, flagVALUE, L"a two"});
		}

		TEST_METHOD(Test_AppendFlagIfNotDupe)
		{
			auto flagArray = std::vector<FLAG_ARRAY_ENTRY>{};
			// Do a stable sort on ulFlagName (first member)
			addin::AppendFlagIfNotDupe(flagArray, {1, 2, flagVALUE, L"two"});
			testFA(L"add 1", flagArray[0], {1, 2, flagVALUE, L"two"});

			addin::AppendFlagIfNotDupe(flagArray, {1, 1, flagVALUE, L"one"});
			testFA(L"add 2", flagArray[1], {1, 1, flagVALUE, L"one"});
			addin::AppendFlagIfNotDupe(flagArray, {1, 1, flagVALUE, L"one"});
			Assert::AreEqual(size_t{2}, flagArray.size(), L"no dupe");

			addin::AppendFlagIfNotDupe(flagArray, {2, 1, flagVALUE, L"b one"});
			testFA(L"add 3", flagArray[2], {2, 1, flagVALUE, L"b one"});
		}
	};
} // namespace addintest