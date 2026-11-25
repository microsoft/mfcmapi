#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/RuleAction.h>

namespace smartview
{
	// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxorule/81b16a56-c4db-42a4-8e51-10e38bf055a6
	class ClientActions : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::vector<std::shared_ptr<RuleAction>> actions;
		std::shared_ptr<blockBytes> remaining = emptyBB();
		bool invalid{};
	};
} // namespace smartview
