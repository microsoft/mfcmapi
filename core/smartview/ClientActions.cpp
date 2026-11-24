#include <core/stdafx.h>
#include <core/smartview/ClientActions.h>

namespace smartview
{
	void ClientActions::parse()
	{
		while (parser->getSize() && actions.size() < _MaxEntriesSmall)
		{
			const auto action = std::make_shared<RuleAction>(false);
			action->block::parse(parser, false);
			actions.push_back(action);
			if (!action->isValid())
			{
				invalid = true;
				break;
			}
		}

		if (parser->getSize())
		{
			invalid = true;
			remaining = blockBytes::parse(parser, parser->getSize());
		}
	}

	void ClientActions::parseBlocks()
	{
		setText(L"Client Actions");
		auto i = 0;
		for (const auto& action : actions)
		{
			addChild(action, L"RuleAction[%1!d!]", i++);
		}

		if (invalid) addHeader(L"Invalid or truncated rule action");
		addLabeledChild(L"Unparsed client actions", remaining);
	}
} // namespace smartview
