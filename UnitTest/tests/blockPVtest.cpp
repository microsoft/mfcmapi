#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/smartview/block/blockPV.h>
#include <core/utility/strings.h>

namespace blockPVtest
{
	TEST_CLASS(blockPVtest)
	{
	private:
		struct pvTestCase
		{
			std::wstring testName;
			ULONG ulPropTag;
			std::wstring source{};
			bool doNickname;
			bool doRuleProcessing;
			size_t offset;
			size_t size;
			std::wstring text;
		};

		void testPV(const pvTestCase& data)
		{
			auto block = smartview::getPVParser(PROP_TYPE(data.ulPropTag));
			auto parser = makeParser(data.source);

			block->parse(parser, data.ulPropTag, data.doNickname, data.doRuleProcessing);
			Assert::AreEqual(data.offset, block->getOffset(), (data.testName + L"-offset").c_str());
			Assert::AreEqual(data.size, block->getSize(), (data.testName + L"-size").c_str());
			unittest::AreEqualEx(data.text.c_str(), block->PropBlock()->c_str(), (data.testName + L"-text").c_str());
		}

		void testPV(const std::vector<pvTestCase>& testCases)
		{
			for (const auto test : testCases)
			{
				testPV(test);
			}
		}

		std::shared_ptr<smartview::binaryParser> makeParser(const std::wstring& str)
		{
			return std::make_shared<smartview::binaryParser>(strings::HexStringToBin(str));
		}

	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_PT_BINARY)
		{
			testPV(std::vector<pvTestCase>{
				pvTestCase{L"bin-f-t",
						   PR_CHANGE_KEY,
						   L"14000000516B013F8BAD5B4D9A74CAB5B37B588400000006",
						   false,
						   true,
						   0,
						   0x18,
						   L"cb: 20 lpb: 516B013F8BAD5B4D9A74CAB5B37B588400000006"},
				pvTestCase{L"bin-f-f",
						   PR_CHANGE_KEY,
						   L"1400516B013F8BAD5B4D9A74CAB5B37B588400000006",
						   false,
						   false,
						   0,
						   0x16,
						   L"cb: 20 lpb: 516B013F8BAD5B4D9A74CAB5B37B588400000006"},
				pvTestCase{L"bin-t-f",
						   PR_CHANGE_KEY,
						   L"000000000000000014000000516B013F8BAD5B4D9A74CAB5B37B588400000006",
						   true,
						   false,
						   0,
						   0x20,
						   L"cb: 20 lpb: 516B013F8BAD5B4D9A74CAB5B37B588400000006"}
			});
		}
	};
} // namespace blockPVtest