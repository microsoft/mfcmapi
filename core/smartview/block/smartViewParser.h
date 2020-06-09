#pragma once
#include <core/smartview/block/block.h>

namespace smartview
{
	class smartViewParser : public block
	{
	public:
		smartViewParser() = default;
		virtual ~smartViewParser() = default;
		smartViewParser(const smartViewParser&) = delete;
		smartViewParser& operator=(const smartViewParser&) = delete;
	};
} // namespace smartview