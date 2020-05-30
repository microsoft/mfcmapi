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
			auto block1 = std::make_shared<smartview::block>(L"test");
			auto block2 = std::make_shared<smartview::block>(L"block2");
			auto block3 = std::make_shared<smartview::block>(L"");
			Assert::AreEqual(block1->isHeader(), true);
			block1->setSize(5);
			Assert::AreEqual(block1->isHeader(), false);
			block1->setOffset(6);
			Assert::AreEqual(block1->isHeader(), false);
			block1->setSource(7);

			Assert::AreEqual(block1->getText(), std::wstring(L"test"));
			Assert::AreEqual(block1->getChildren().empty(), true);
			Assert::AreEqual(block1->toString(), std::wstring(L"test"));
			Assert::AreEqual(block1->getSize(), size_t(5));
			Assert::AreEqual(block1->getOffset(), size_t(6));
			Assert::AreEqual(block1->getSource(), ULONG(7));

			block1->setText(L"the %1!ws!", L"other");
			Assert::AreEqual(block1->getText(), std::wstring(L"the other"));
			block1->addHeader(L" this %1!ws!", L"that");
			Assert::AreEqual(block1->toString(), std::wstring(L"the other this that"));

			block1->setText(L"hello");
			Assert::AreEqual(block1->getText(), std::wstring(L"hello"));

			block1->setText(L"hello %1!ws!", L"world");
			Assert::AreEqual(block1->getText(), std::wstring(L"hello world"));

			block1->terminateBlock();
			block1->terminateBlock();
			block1->addChild(block2);
			Assert::AreEqual(block1->toString(), std::wstring(L"hello world this that\r\nblock2"));

			block1->addBlankLine();
			block1->addLabeledChild(L"Label: ", block2);
			Assert::AreEqual(
				block1->toString(), std::wstring(L"hello world this that\r\nblock2\r\n\r\nLabel: block2\r\n"));

			Assert::AreEqual(block3->hasData(), false);
			block3->addChild(block1);
			Assert::AreEqual(block3->hasData(), true);
			block3->setSource(42);
			Assert::AreEqual(block3->getSource(), ULONG(42));
			Assert::AreEqual(block1->getSource(), ULONG(42));
		}

		TEST_METHOD(Test_blockStringA)
		{
			auto block1 = std::make_shared<smartview::blockStringA>();
			Assert::AreEqual(block1->isSet(), false);
		}

		TEST_METHOD(Test_blockStringW)
		{
			auto block1 = std::make_shared<smartview::blockStringW>();
			Assert::AreEqual(block1->isSet(), false);
			auto block2 = std::make_shared<smartview::blockStringW>(std::wstring(L"test"), 4, 5);
			Assert::AreEqual(block2->length(), size_t(4));
		}
	};
} // namespace blocktest