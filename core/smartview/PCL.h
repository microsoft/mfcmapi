#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct SizedXID
	{
		std::shared_ptr<blockT<BYTE>> XidSize = emptyT<BYTE>();
		std::shared_ptr<blockT<GUID>> NamespaceGuid = emptyT<GUID>();
		std::shared_ptr<blockBytes> LocalID = emptyBB();

		SizedXID(std::shared_ptr<binaryParser>& parser);
	};

	class PCL : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::vector<std::shared_ptr<SizedXID>> m_lpXID;
	};
} // namespace smartview