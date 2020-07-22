#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class SizedXID : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> XidSize = emptyT<BYTE>();
		std::shared_ptr<blockT<GUID>> NamespaceGuid = emptyT<GUID>();
		std::shared_ptr<blockBytes> LocalID = emptyBB();
	};

	class PCL : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::vector<std::shared_ptr<SizedXID>> m_lpXID;
	};
} // namespace smartview