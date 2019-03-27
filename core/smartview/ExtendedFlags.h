#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct ExtendedFlag
	{
		std::shared_ptr<blockT<BYTE>> Id = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Cb = emptyT<BYTE>();
		struct
		{
			std::shared_ptr<blockT<DWORD>> ExtendedFlags = emptyT<DWORD>();
			std::shared_ptr<blockT<GUID>> SearchFolderID = emptyT<GUID>();
			std::shared_ptr<blockT<DWORD>> SearchFolderTag = emptyT<DWORD>();
			std::shared_ptr<blockT<DWORD>> ToDoFolderVersion = emptyT<DWORD>();
		} Data;
		std::shared_ptr<blockBytes> lpUnknownData = emptyBB();
		bool bBadData{};

		ExtendedFlag(const std::shared_ptr<binaryParser>& parser);
	};

	class ExtendedFlags : public smartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::vector<std::shared_ptr<ExtendedFlag>> m_pefExtendedFlags;
	};
} // namespace smartview