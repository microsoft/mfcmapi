#pragma once
#include <winnt.h>
#include <core/smartview/SmartViewParser.h>

namespace smartview
{
	struct SizedXID
	{
		blockT<BYTE> XidSize;
		blockT<GUID> NamespaceGuid;
		DWORD cbLocalId{};
		blockBytes LocalID;
	};

	class PCL : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		DWORD m_cXID{};
		std::vector<SizedXID> m_lpXID;
	};
} // namespace smartview