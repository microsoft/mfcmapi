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
			size_t offset,
			size_t size)
		{
			auto block = smartview::getPVParser(ulPropTag, doNickname, doRuleProcessing);
			auto parser = makeParser(source);

			block->block::parse(parser, false);
			Assert::AreEqual(offset, block->getOffset(), (testName + L"-offset").c_str());
			Assert::AreEqual(size, block->getSize(), (testName + L"-size").c_str());
			unittest::AreEqualEx(propblock.c_str(), block->toString(), (testName + L"-propblock").c_str());
			// Check that we consumed the entire input
			Assert::AreEqual(true, parser->empty(), (testName + L"-complete").c_str());
		}

	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_PT_BINARY)
		{
			auto propblock = std::wstring(L"PropString = cb: 20 lpb: 516B013F8BAD5B4D9A74CAB5B37B588400000006\r\n"
										  L"AltPropString = Qk.?.≠[M.t µ≥{X.....\r\n"
										  L"Smart View\r\n"
										  L"\tXID\r\n"
										  L"\t\tNamespaceGuid = {3F016B51-AD8B-4D5B-9A74-CAB5B37B5884}\r\n"
										  L"\t\tLocalId = cb: 4 lpb: 00000006");
			testPV(
				L"bin-f-f",
				PR_CHANGE_KEY,
				L"1400516B013F8BAD5B4D9A74CAB5B37B588400000006",
				false,
				false,
				propblock,
				0,
				0x16);
			testPV(
				L"bin-f-t",
				PR_CHANGE_KEY,
				L"14000000516B013F8BAD5B4D9A74CAB5B37B588400000006",
				false,
				true,
				propblock,
				0,
				0x18);
			testPV(
				L"bin-t-f",
				PR_CHANGE_KEY,
				L"000000000000000014000000516B013F8BAD5B4D9A74CAB5B37B588400000006",
				true,
				false,
				propblock,
				0,
				0x20);
		}

		TEST_METHOD(Test_PT_UNICODE)
		{
			auto propblock = std::wstring(L"PropString = test string\r\n"
										  L"AltPropString = cb: 22 lpb: 7400650073007400200073007400720069006E006700");
			auto smartview = std::wstring(L"");
			testPV(
				L"uni-f-f",
				PR_SUBJECT_W,
				L"16007400650073007400200073007400720069006E006700",
				false,
				false,
				propblock,
				0,
				0x18);
			testPV(
				L"uni-f-t",
				PR_SUBJECT_W,
				L"7400650073007400200073007400720069006E0067000000",
				false,
				true,
				propblock,
				0,
				0x18);
			testPV(
				L"uni-t-f",
				PR_SUBJECT_W,
				L"0000000000000000160000007400650073007400200073007400720069006E006700",
				true,
				false,
				propblock,
				0,
				0x22);
		}

		TEST_METHOD(Test_PT_LONG)
		{
			auto propblock = std::wstring(L"PropString = 9\r\n"
										  L"AltPropString = 0x9\r\n"
										  L"Smart View\r\n"
										  L"\tFlags: MAPI_PROFSECT");
			testPV(L"long-f-f", PR_OBJECT_TYPE, L"09000000", false, false, propblock, 0, 0x4);
			testPV(L"long-f-t", PR_OBJECT_TYPE, L"09000000", false, true, propblock, 0, 0x4);
			testPV(L"long-t-f", PR_OBJECT_TYPE, L"0900000000000000", true, false, propblock, 0, 0x8);
		}

		TEST_METHOD(Test_PT_SYSTIME)
		{
			auto propblock = std::wstring(L"PropString = 03:59:00.000 AM 5/10/2020\r\n"
										  L"AltPropString = Low: 0x5606DA00 High: 0x01D6267F");
			testPV(L"systime-f-f", PR_MESSAGE_DELIVERY_TIME, L"00DA06567F26D601", false, false, propblock, 0, 0x8);
			testPV(L"systime-f-t", PR_MESSAGE_DELIVERY_TIME, L"00DA06567F26D601", false, true, propblock, 0, 0x8);
			testPV(L"systime-t-f", PR_MESSAGE_DELIVERY_TIME, L"00DA06567F26D601", true, false, propblock, 0, 0x8);
		}

		TEST_METHOD(Test_PT_BOOLEAN)
		{
			testPV(L"bool-f-f-1", PR_MESSAGE_TO_ME, L"0100", false, false, L"PropString = True", 0, 0x2);
			testPV(L"bool-f-f" - 2, PR_MESSAGE_TO_ME, L"0000", false, false, L"PropString = False", 0, 0x2);
			testPV(L"bool-f-t", PR_MESSAGE_TO_ME, L"01", false, true, L"PropString = True", 0, 0x1);
			testPV(L"bool-t-f", PR_MESSAGE_TO_ME, L"0100000000000000", true, false, L"PropString = True", 0, 0x8);
		}
	};
} // namespace blockPVtest