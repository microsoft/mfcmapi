#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	struct ExtendedFlag
	{
		BYTE Id{};
		BYTE Cb{};
		union
		{
			DWORD ExtendedFlags;
			GUID SearchFolderID;
			DWORD SearchFolderTag;
			DWORD ToDoFolderVersion;
		} Data{};
		std::vector<BYTE> lpUnknownData;
	};

	class ExtendedFlags : public SmartViewParser
	{
	public:
		ExtendedFlags();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		ULONG m_ulNumFlags;
		std::vector<ExtendedFlag> m_pefExtendedFlags;
	};
}