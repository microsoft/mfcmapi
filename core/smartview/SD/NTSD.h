#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/interpret/sid.h>

namespace smartview
{
	class NTSD : public block
	{
	public:
		NTSD(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB);

	private:
		void parse() override;
		void parseBlocks() override;

		sid::aceType acetype{sid::aceType::Message};
		std::shared_ptr<blockBytes> m_SDbin = emptyBB();
	};
} // namespace smartview