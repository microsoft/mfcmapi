#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

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
		std::shared_ptr<blockBytes> lpUnknownData = emptyBB();
		bool bBadData{};

		ExtendedFlag(std::shared_ptr<binaryParser> parser);
	};

	class ExtendedFlags : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::vector<std::shared_ptr<ExtendedFlag>> m_pefExtendedFlags;
	};
} // namespace smartview