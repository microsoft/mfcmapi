#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	struct ExtendedFlag
	{
		blockT<BYTE> Id;
		blockT<BYTE> Cb;
		struct
		{
			blockT<DWORD> ExtendedFlags;
			blockT<GUID> SearchFolderID;
			blockT<DWORD> SearchFolderTag;
			blockT<DWORD> ToDoFolderVersion;
		} Data;
		blockBytes lpUnknownData;
	};

	class ExtendedFlags : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		ULONG m_ulNumFlags{};
		std::vector<ExtendedFlag> m_pefExtendedFlags;
	};
}