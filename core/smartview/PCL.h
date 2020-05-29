#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct SizedXID
	{
		std::shared_ptr<blockT<BYTE>> XidSize = emptyT<BYTE>();
		std::shared_ptr<blockT<GUID>> NamespaceGuid = emptyT<GUID>();
		std::shared_ptr<blockBytes> LocalID = emptyBB();

		SizedXID(const std::shared_ptr<binaryParser>& parser);
	};

	class PCL : public smartViewParser
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::vector<std::shared_ptr<SizedXID>> m_lpXID;
	};
} // namespace smartview