#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/addin/mfcmapi.h>

namespace smartview
{
	class addinParser : public block
	{
	public:
		addinParser(parserType _type) : type{_type} {}

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockBytes> bin = emptyBB();
		parserType type;
	};
} // namespace smartview