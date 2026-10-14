#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/addin/addin.h>
#include <core/addin/addinfile.h>
#include <core/addin/mfcmapi.h>
#include <cstring>

namespace addin
{
	// imports for testing
	int _cdecl compareTypes(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareTags(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareNameID(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept;
	int _cdecl compareSmartViewParser(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept;

	std::wstring AddInStructTypeToString(parserType parser);

	void SortFlagArray(_In_count_(ulFlags) LPFLAG_ARRAY_ENTRY lpFlags, _In_ size_t ulFlags) noexcept;
	void AppendFlagIfNotDupe(std::vector<FLAG_ARRAY_ENTRY>& target, FLAG_ARRAY_ENTRY source);

	void MergeFlagArrays(std::vector<FLAG_ARRAY_ENTRY>& In1, _In_count_(cIn2) LPFLAG_ARRAY_ENTRY In2, _In_ size_t cIn2);
} // namespace addin

namespace addintest
{
	int _cdecl testCompareTypes(_In_ const NAME_ARRAY_ENTRY& a1, _In_ const NAME_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareTypes(nullptr, reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareTags(_In_ const NAME_ARRAY_ENTRY_V2& a1, _In_ const NAME_ARRAY_ENTRY_V2& a2) noexcept
	{
		return addin::compareTags(nullptr, reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareNameID(_In_ const NAMEID_ARRAY_ENTRY& a1, _In_ const NAMEID_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareNameID(nullptr, reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	int _cdecl testCompareSmartViewParser(
		_In_ const SMARTVIEW_PARSER_ARRAY_ENTRY& a1,
		_In_ const SMARTVIEW_PARSER_ARRAY_ENTRY& a2) noexcept
	{
		return addin::compareSmartViewParser(nullptr, reinterpret_cast<LPCVOID>(&a1), reinterpret_cast<LPCVOID>(&a2));
	}

	std::vector<BYTE> MakePeImageWithExport(const char* exportName, bool is64Bit = true)
	{
		constexpr size_t ntOffset = 0x80;
		constexpr size_t rawOffset = 0x200;
		auto image = std::vector<BYTE>(0x400);

		IMAGE_DOS_HEADER dosHeader = {};
		dosHeader.e_magic = IMAGE_DOS_SIGNATURE;
		dosHeader.e_lfanew = static_cast<LONG>(ntOffset);
		memcpy(image.data(), &dosHeader, sizeof(dosHeader));

		const DWORD signature = IMAGE_NT_SIGNATURE;
		memcpy(image.data() + ntOffset, &signature, sizeof(signature));

		IMAGE_FILE_HEADER fileHeader = {};
		fileHeader.Machine = is64Bit ? IMAGE_FILE_MACHINE_AMD64 : IMAGE_FILE_MACHINE_I386;
		fileHeader.NumberOfSections = 1;
		fileHeader.SizeOfOptionalHeader =
			static_cast<WORD>(is64Bit ? sizeof(IMAGE_OPTIONAL_HEADER64) : sizeof(IMAGE_OPTIONAL_HEADER32));
		memcpy(image.data() + ntOffset + sizeof(signature), &fileHeader, sizeof(fileHeader));

		const auto optionalOffset = ntOffset + sizeof(signature) + sizeof(fileHeader);
		if (is64Bit)
		{
			IMAGE_OPTIONAL_HEADER64 optionalHeader = {};
			optionalHeader.Magic = IMAGE_NT_OPTIONAL_HDR64_MAGIC;
			optionalHeader.SizeOfHeaders = static_cast<DWORD>(rawOffset);
			optionalHeader.NumberOfRvaAndSizes = IMAGE_NUMBEROF_DIRECTORY_ENTRIES;
			optionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT] =
				IMAGE_DATA_DIRECTORY{0x1000, sizeof(IMAGE_EXPORT_DIRECTORY)};
			memcpy(image.data() + optionalOffset, &optionalHeader, sizeof(optionalHeader));
		}
		else
		{
			IMAGE_OPTIONAL_HEADER32 optionalHeader = {};
			optionalHeader.Magic = IMAGE_NT_OPTIONAL_HDR32_MAGIC;
			optionalHeader.SizeOfHeaders = static_cast<DWORD>(rawOffset);
			optionalHeader.NumberOfRvaAndSizes = IMAGE_NUMBEROF_DIRECTORY_ENTRIES;
			optionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT] =
				IMAGE_DATA_DIRECTORY{0x1000, sizeof(IMAGE_EXPORT_DIRECTORY)};
			memcpy(image.data() + optionalOffset, &optionalHeader, sizeof(optionalHeader));
		}

		IMAGE_SECTION_HEADER section = {};
		section.VirtualAddress = 0x1000;
		section.SizeOfRawData = 0x200;
		section.PointerToRawData = static_cast<DWORD>(rawOffset);
		const auto sectionOffset = optionalOffset + fileHeader.SizeOfOptionalHeader;
		memcpy(image.data() + sectionOffset, &section, sizeof(section));

		IMAGE_EXPORT_DIRECTORY exports = {};
		exports.NumberOfNames = 1;
		exports.AddressOfNames = 0x1040;
		memcpy(image.data() + rawOffset, &exports, sizeof(exports));

		const DWORD nameRva = 0x1060;
		memcpy(image.data() + rawOffset + 0x40, &nameRva, sizeof(nameRva));
		memcpy(image.data() + rawOffset + 0x60, exportName, strlen(exportName) + 1);
		return image;
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

		//AddInStructTypeToString
		TEST_METHOD(Test_AddInStructTypeToString)
		{
			Assert::AreEqual(
				L"Choose Smart View Parser",
				addin::AddInStructTypeToString(parserType::NOPARSING).c_str(),
				L"NOPARSING");
			Assert::AreEqual(L"", addin::AddInStructTypeToString(parserType{0xfff}).c_str(), L"unknown");
		}

		TEST_METHOD(Test_HasLoadAddInExport)
		{
			auto addInImage = MakePeImageWithExport("LoadAddIn");
			Assert::IsTrue(addin::HasLoadAddInExport(addInImage.data(), addInImage.size()));
			Assert::IsTrue(addin::IsCompatibleBitness(addInImage.data(), addInImage.size(), true));
			Assert::IsFalse(addin::IsCompatibleBitness(addInImage.data(), addInImage.size(), false));

			auto x86Image = MakePeImageWithExport("LoadAddIn", false);
			Assert::IsTrue(addin::HasLoadAddInExport(x86Image.data(), x86Image.size()));
			Assert::IsTrue(addin::IsCompatibleBitness(x86Image.data(), x86Image.size(), false));
			Assert::IsFalse(addin::IsCompatibleBitness(x86Image.data(), x86Image.size(), true));

			auto otherImage = MakePeImageWithExport("NotAnAddIn");
			Assert::IsFalse(addin::HasLoadAddInExport(otherImage.data(), otherImage.size()));

			auto invalidNameRva = MakePeImageWithExport("LoadAddIn");
			const DWORD invalidRva = 0xfffffff0;
			memcpy(invalidNameRva.data() + 0x240, &invalidRva, sizeof(invalidRva));
			Assert::IsFalse(addin::HasLoadAddInExport(invalidNameRva.data(), invalidNameRva.size()));

			auto unterminatedName = MakePeImageWithExport("LoadAddIn");
			memset(unterminatedName.data() + 0x260, 'A', unterminatedName.size() - 0x260);
			Assert::IsFalse(addin::HasLoadAddInExport(unterminatedName.data(), unterminatedName.size()));

			addInImage.resize(32);
			Assert::IsFalse(addin::HasLoadAddInExport(addInImage.data(), addInImage.size()));
		}

		void testFA(std::wstring testName, const FLAG_ARRAY_ENTRY& fa1, const FLAG_ARRAY_ENTRY& fa2)
		{
			Assert::AreEqual(fa1.ulFlagName, fa2.ulFlagName, (testName + L"-ulFlagName").c_str());
			Assert::AreEqual(fa1.lFlagValue, fa2.lFlagValue, (testName + L"-lFlagValue").c_str());
			Assert::AreEqual(fa1.ulFlagType, fa2.ulFlagType, (testName + L"-ulFlagType").c_str());
			Assert::AreEqual(fa1.lpszName, fa2.lpszName, (testName + L"-lpszName").c_str());
		}

		TEST_METHOD(Test_FlagArray)
		{
			FLAG_ARRAY_ENTRY flagArray[] = {
				{2, 1, flagVALUE, L"b one"},
				{1, 2, flagVALUE, L"two"},
				{2, 3, flagVALUE, L"a three"},
				{1, 1, flagVALUE, L"one"},
			};

			// Do a stable sort on ulFlagName (first member)
			addin::SortFlagArray(flagArray, _countof(flagArray));

			testFA(L"0", flagArray[0], {1, 2, flagVALUE, L"two"});
			testFA(L"1", flagArray[1], {1, 1, flagVALUE, L"one"});
			testFA(L"2", flagArray[2], {2, 1, flagVALUE, L"b one"});
			testFA(L"3", flagArray[3], {2, 3, flagVALUE, L"a three"});

			auto flagV = std::vector<FLAG_ARRAY_ENTRY>{};
			addin::AppendFlagIfNotDupe(flagV, {1, 3, flagVALUE, L"three"});
			testFA(L"add 1", flagV[0], {1, 3, flagVALUE, L"three"});

			addin::AppendFlagIfNotDupe(flagV, {1, 4, flagVALUE, L"one"});
			testFA(L"add 2", flagV[1], {1, 4, flagVALUE, L"one"});
			addin::AppendFlagIfNotDupe(flagV, {1, 4, flagVALUE, L"one"});
			Assert::AreEqual(size_t{2}, flagV.size(), L"no dupe");

			addin::AppendFlagIfNotDupe(flagV, {2, 2, flagVALUE, L"c two"});
			testFA(L"add 3", flagV[2], {2, 2, flagVALUE, L"c two"});

			addin::MergeFlagArrays(flagV, flagArray, _countof(flagArray));
			Assert::AreEqual(size_t{7}, flagV.size(), L"merge size");

			testFA(L"m0", flagV[0], {1, 3, flagVALUE, L"three"});
			testFA(L"m1", flagV[1], {1, 4, flagVALUE, L"one"});
			testFA(L"m2", flagV[2], {1, 2, flagVALUE, L"two"});
			testFA(L"m3", flagV[3], {1, 1, flagVALUE, L"one"});
			testFA(L"m4", flagV[4], {2, 2, flagVALUE, L"c two"});
			testFA(L"m5", flagV[5], {2, 1, flagVALUE, L"b one"});
			testFA(L"m6", flagV[6], {2, 3, flagVALUE, L"a three"});

			auto flagV2 = std::vector<FLAG_ARRAY_ENTRY>{};
			addin::AppendFlagIfNotDupe(flagV2, {4, 3, flagVALUE, L"three"});
			testFA(L"f2 add 1", flagV2[0], {4, 3, flagVALUE, L"three"});
			addin::MergeFlagArrays(flagV2, nullptr, 0);
			Assert::AreEqual(size_t{1}, flagV2.size(), L"merge null");

			addin::AppendFlagIfNotDupe(flagV2, {4, 4, flagVALUE, L"four"});
			Assert::AreEqual(size_t{2}, flagV2.size(), L"merge null");
			testFA(L"f2 add 2", flagV2[1], {4, 4, flagVALUE, L"four"});

			FLAG_ARRAY_ENTRY flagArray2[] = {{3, 1, flagVALUE, L"four one"}};
			addin::MergeFlagArrays(flagV2, flagArray2, _countof(flagArray2));
			Assert::AreEqual(size_t{3}, flagV2.size(), L"merge null");
			testFA(L"f2 merge", flagV2[0], {3, 1, flagVALUE, L"four one"});
		}
	};
} // namespace addintest