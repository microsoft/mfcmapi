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

	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(Test_PT_BINARY)
		{
			auto block1 = std::make_shared<smartview::SBinaryBlock>();
			auto bin1 =
				makeParser(L"250000004449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D");
			block1->parse(bin1, PR_SENDER_SEARCH_KEY);
			Assert::AreEqual(size_t{0}, block1->PropBlock()->getOffset());
			Assert::AreEqual(size_t{0x29}, block1->PropBlock()->getSize());
			Assert::AreEqual(
				L"cb: 37 lpb: 4449534E45595641434154494F4E434C5542404D41494C2E4456434D454D4245522E434F4D",
				block1->PropBlock()->c_str());
		}
	};
} // namespace blockPVtest