#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/addin/mfcmapi.h>
#include <core/smartview/RuleAction.h>
#include <core/smartview/SmartView.h>
#include <core/utility/strings.h>

namespace ruleactiontest
{
	TEST_CLASS(ruleactiontest)
	{
	public:
		static const bool dummy_var = true;
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD(ClientActionsPropertyUsesClientActionsParser)
		{
			const auto type =
				smartview::FindSmartViewParserForProp(PR_CLIENT_ACTIONS, nullptr, nullptr, nullptr, false, false);
			Assert::IsTrue(type == parserType::CLIENTACTIONS);
			Assert::IsNotNull(smartview::GetSmartViewParser(type, nullptr).get());
		}

		TEST_METHOD(RuleActionRejectsMissingActionBlocks)
		{
			// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/a31c67b8-405e-4c4f-bfda-e5be0091041a
			for (const auto& test :
				 {std::pair{false, L"020009000B0000000000000000FF"},
				  std::pair{true, L"00000100000002000000090000000B0000000000000000FF"}})
			{
				const auto parser = std::make_shared<smartview::binaryParser>(strings::HexStringToBin(test.second));
				auto action = smartview::RuleAction(test.first);
				action.block::parse(parser, false);

				Assert::IsFalse(action.isValid());
				Assert::IsTrue(parser->empty());
			}
		}

		TEST_METHOD(RuleActionRejectsInvalidAndTruncatedActionCount)
		{
			// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/a31c67b8-405e-4c4f-bfda-e5be0091041a
			for (const auto bytes : {L"0000", L"FF"})
			{
				const auto parser = std::make_shared<smartview::binaryParser>(strings::HexStringToBin(bytes));
				auto action = smartview::RuleAction(false);
				action.block::parse(parser, false);

				Assert::IsFalse(action.isValid());
			}
		}
	};
} // namespace ruleactiontest
