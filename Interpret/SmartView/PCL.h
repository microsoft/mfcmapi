#pragma once
#include <winnt.h>
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	struct SizedXID
	{
		BYTE XidSize{};
		GUID NamespaceGuid{};
		DWORD cbLocalId{};
		std::vector<BYTE> LocalID;
	};

	class PCL : public SmartViewParser
	{
	public:
		PCL();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		DWORD m_cXID;
		std::vector<SizedXID> m_lpXID;
	};
}