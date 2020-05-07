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
			ULONG ulPropTag;
			std::wstring source{};
			size_t offset;
			size_t size;
			std::wstring text;
			bool doNickname;
			bool doRuleProcessing;
		};

		void testPV(const pvTestCase & data)
		{
			auto block = smartview::getPVParser(PROP_TYPE(data.ulPropTag));
			auto parser = makeParser(data.source);

			block->parse(parser, data.ulPropTag, data.doNickname, data.doRuleProcessing);
			Assert::AreEqual(data.offset, block->getOffset());
			Assert::AreEqual(data.size, block->getSize());
			Assert::AreEqual(data.text.c_str(), block->PropBlock()->c_str());
		}

		void testPV(const std::vector<pvTestCase> & testCases)
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
				pvTestCase{PR_SENDER_SEARCH_KEY,
						   L"250000004449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D",
						   0,
						   0x29,
						   L"cb: 37 lpb: 4449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D",
						   false,
						   true},
				pvTestCase{PR_SENDER_SEARCH_KEY,
						   L"25004449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D",
						   0,
						   0x27,
						   L"cb: 37 lpb: 4449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D",
						   false,
						   false},
				pvTestCase{PR_SENDER_SEARCH_KEY,
						   L"0000000000000000250000004449534E45595641434154494F4E434C5542404D41494C2E4456434D454D424552"
						   L"2E434F4D",
						   0,
						   0x31,
						   L"cb: 37 lpb: 4449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D",
						   true,
						   false}});
		}
	};
} // namespace blockPVtest