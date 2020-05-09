#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/smartview/block/blockPV.h>
#include <core/utility/strings.h>

namespace blockPVtest
{
	TEST_CLASS(blockPVtest)
	{
	private:
		std::shared_ptr<smartview::binaryParser> makeParser(const std::wstring& str)
		{
			return std::make_shared<smartview::binaryParser>(strings::HexStringToBin(str));
		}

		void testPV(
			std::wstring testName,
			ULONG ulPropTag,
			std::wstring source,
			bool doNickname,
			bool doRuleProcessing,
			std::wstring propblock,
			std::wstring altpropblock,
			std::wstring smartview,
			size_t offset,
			size_t size)
		{
			auto block = smartview::getPVParser(PROP_TYPE(ulPropTag));
			auto parser = makeParser(source);

			block->parse(parser, ulPropTag, doNickname, doRuleProcessing);
			Assert::AreEqual(offset, block->getOffset(), (testName + L"-offset").c_str());
			Assert::AreEqual(size, block->getSize(), (testName + L"-size").c_str());
			unittest::AreEqualEx(propblock.c_str(), block->PropBlock()->c_str(), (testName + L"-propblock").c_str());
			unittest::AreEqualEx(
				altpropblock.c_str(), block->AltPropBlock()->c_str(), (testName + L"-altpropblock").c_str());
			unittest::AreEqualEx(
				smartview.c_str(), block->SmartViewBlock()->c_str(), (testName + L"-smartview").c_str());
			// Check that we consumed the entire input
			Assert::AreEqual(true, parser->empty(), (testName + L"-complete").c_str());
			// overflow
			Assert::AreEqual(false, parser->overflow(), (testName + L"-no-overflow").c_str());
		}

	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_PT_BINARY)
		{
			auto propblock = std::wstring(L"cb: 20 lpb: 516B013F8BAD5B4D9A74CAB5B37B588400000006");
			auto altpropblock = std::wstring(L"Qk.?.≠[M.t µ≥{X.....");
			auto smartview = std::wstring(L"XID:\r\n"
										  L"NamespaceGuid = {3F016B51-AD8B-4D5B-9A74-CAB5B37B5884}\r\n"
										  L"LocalId = cb: 4 lpb: 00000006");
			testPV(
				L"bin-f-t",
				PR_CHANGE_KEY,
				L"14000000516B013F8BAD5B4D9A74CAB5B37B588400000006",
				false,
				true,
				propblock,
				altpropblock,
				smartview,
				0,
				0x18);
			testPV(
				L"bin-f-f",
				PR_CHANGE_KEY,
				L"1400516B013F8BAD5B4D9A74CAB5B37B588400000006",
				false,
				false,
				propblock,
				altpropblock,
				smartview,
				0,
				0x16);
			testPV(
				L"bin-t-f",
				PR_CHANGE_KEY,
				L"000000000000000014000000516B013F8BAD5B4D9A74CAB5B37B588400000006",
				true,
				false,
				propblock,
				altpropblock,
				smartview,
				0,
				0x20);
		}

		TEST_METHOD(Test_PT_UNICODE)
		{
			auto propblock = std::wstring(L"test string");
			auto altpropblock = std::wstring(L"cb: 22 lpb: 7400650073007400200073007400720069006E006700");
			auto smartview = std::wstring(L"");
			testPV(
				L"uni-f-t",
				PR_SUBJECT_W,
				L"7400650073007400200073007400720069006E0067000000",
				false,
				true,
				propblock,
				altpropblock,
				smartview,
				0,
				0x18);
			testPV(
				L"uni-f-f",
				PR_SUBJECT_W,
				L"16007400650073007400200073007400720069006E006700",
				false,
				false,
				propblock,
				altpropblock,
				smartview,
				0,
				0x18);
			testPV(
				L"uni-t-f",
				PR_SUBJECT_W,
				L"0000000000000000160000007400650073007400200073007400720069006E006700",
				true,
				false,
				propblock,
				altpropblock,
				smartview,
				0,
				0x22);
		}

		TEST_METHOD(Test_PT_LONG)
		{
			auto propblock = std::wstring(L"9");
			auto altpropblock = std::wstring(L"0x9");
			auto smartview = std::wstring(L"Flags: MAPI_PROFSECT");
			testPV(L"long-f-t", PR_OBJECT_TYPE, L"09000000", false, true, propblock, altpropblock, smartview, 0, 0x4);
			testPV(L"long-f-f", PR_OBJECT_TYPE, L"09000000", false, false, propblock, altpropblock, smartview, 0, 0x4);
			testPV(
				L"long-t-f",
				PR_OBJECT_TYPE,
				L"0900000000000000",
				true,
				false,
				propblock,
				altpropblock,
				smartview,
				0,
				0x8);
		}
	};
} // namespace blockPVtest