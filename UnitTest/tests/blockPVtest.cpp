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
			std::shared_ptr<smartview::blockPV> block{};
			ULONG ulPropTag;
			std::shared_ptr<smartview::binaryParser> parser{};
			size_t offset;
			size_t size;
			std::wstring text;
			bool doNickname;
			bool doRuleProcessing;
		};

		void testPV(pvTestCase & data)
		{
			data.block->parse(data.parser, data.ulPropTag, data.doNickname, data.doRuleProcessing);
			Assert::AreEqual(data.offset, data.block->getOffset());
			Assert::AreEqual(data.size, data.block->getSize());
			Assert::AreEqual(data.text.c_str(), data.block->PropBlock()->c_str());
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
			auto t1 = pvTestCase{
				std::make_shared<smartview::SBinaryBlock>(),
				PR_SENDER_SEARCH_KEY,
				makeParser(L"250000004449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D"),
				0,
				0x29,
				L"cb: 37 lpb: 4449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D",
				false,
				true};
			testPV(t1);
		}
	};
} // namespace blockPVtest