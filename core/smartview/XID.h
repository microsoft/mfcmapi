#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class XID : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::shared_ptr<blockT<GUID>> m_NamespaceGuid = emptyT<GUID>();
		std::shared_ptr<blockBytes> m_LocalID = emptyBB();
	};
} // namespace smartview