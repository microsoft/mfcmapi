#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	struct SizedXID
	{
		blockT<BYTE> XidSize;
		blockT<GUID> NamespaceGuid;
		DWORD cbLocalId{};
		blockBytes LocalID;

		SizedXID(std::shared_ptr<binaryParser>& parser);
	};

	class PCL : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		DWORD m_cXID{};
		std::vector<std::shared_ptr<SizedXID>> m_lpXID;
	};
} // namespace smartview