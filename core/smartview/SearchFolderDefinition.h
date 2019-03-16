#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/EntryList.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct AddressListEntryStruct
	{
		blockT<DWORD> PropertyCount;
		blockT<DWORD> Pad;
		PropertiesStruct Props;

		AddressListEntryStruct(std::shared_ptr<binaryParser> parser);
	};

	class SearchFolderDefinition : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<DWORD> m_Version;
		blockT<DWORD> m_Flags;
		blockT<DWORD> m_NumericSearch;
		blockT<BYTE> m_TextSearchLength;
		blockT<WORD> m_TextSearchLengthExtended;
		std::shared_ptr<blockStringW> m_TextSearch = emptySW();
		blockT<DWORD> m_SkipLen1;
		blockBytes m_SkipBytes1;
		blockT<DWORD> m_DeepSearch;
		blockT<BYTE> m_FolderList1Length;
		blockT<WORD> m_FolderList1LengthExtended;
		std::shared_ptr<blockStringW> m_FolderList1 = emptySW();
		blockT<DWORD> m_FolderList2Length;
		EntryList m_FolderList2;
		blockT<DWORD> m_AddressCount; // SFST_BINARY
		std::vector<std::shared_ptr<AddressListEntryStruct>> m_Addresses; // SFST_BINARY
		blockT<DWORD> m_SkipLen2;
		blockBytes m_SkipBytes2;
		std::shared_ptr<RestrictionStruct> m_Restriction; // SFST_MRES
		blockBytes m_AdvancedSearchBytes; // SFST_FILTERSTREAM
		blockT<DWORD> m_SkipLen3;
		blockBytes m_SkipBytes3;
	};
} // namespace smartview