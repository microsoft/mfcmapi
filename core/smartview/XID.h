#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class XID : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<GUID>> m_NamespaceGuid = emptyT<GUID>();
		std::shared_ptr<blockBytes> m_LocalID = emptyBB();
	};
} // namespace smartview