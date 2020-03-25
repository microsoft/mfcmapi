#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>

namespace blocktest
{
	TEST_CLASS(blocktest)
	{
	public:
		// Without this, clang gets weird
		static const bool dummy_var = true;

		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }
		TEST_METHOD(Test_block)
		{
			auto block1 = smartview::block(L"test");
			Assert::AreEqual(block1.isHeader(), true);
			block1.setSize(5);
			Assert::AreEqual(block1.isHeader(), false);
			block1.setOffset(6);
			Assert::AreEqual(block1.isHeader(), false);
			block1.setSource(7);

			Assert::AreEqual(block1.getText(), std::wstring(L"test"));
			Assert::AreEqual(block1.getChildren().empty(), true);
			Assert::AreEqual(block1.toString(), std::wstring(L"test"));
			Assert::AreEqual(block1.getSize(), size_t(5));
			Assert::AreEqual(block1.getOffset(), size_t(6));
			Assert::AreEqual(block1.getSource(), ULONG(7));

			block1.setText(L"the %1!ws!", L"other");
			Assert::AreEqual(block1.getText(), std::wstring(L"the other"));
			block1.addHeader(L" this %1!ws!", L"that");
			Assert::AreEqual(block1.toString(), std::wstring(L"the other this that"));
		}
	};
} // namespace blocktest