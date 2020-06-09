#pragma once
#include <core/smartview/block/block.h>

namespace smartview
{
	class scratchBlock : public block
	{
	public:
		scratchBlock() noexcept { parsed = true; }

	private:
		void parse() override{};
	};
} // namespace smartview