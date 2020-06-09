#pragma once
#include <core/smartview/block/block.h>
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
		std::shared_ptr<blockT<DWORD>> PropertyCount = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Pad = emptyT<DWORD>();
		std::shared_ptr<PropertiesStruct> Props;

		AddressListEntryStruct(const std::shared_ptr<binaryParser>& parser);
	};

	class SearchFolderDefinition : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_Version = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_Flags = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_NumericSearch = emptyT<DWORD>();
		std::shared_ptr<blockT<BYTE>> m_TextSearchLength = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> m_TextSearchLengthExtended = emptyT<WORD>();
		std::shared_ptr<blockStringW> m_TextSearch = emptySW();
		std::shared_ptr<blockT<DWORD>> m_SkipLen1 = emptyT<DWORD>();
		std::shared_ptr<blockBytes> m_SkipBytes1 = emptyBB();
		std::shared_ptr<blockT<DWORD>> m_DeepSearch = emptyT<DWORD>();
		std::shared_ptr<blockT<BYTE>> m_FolderList1Length = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> m_FolderList1LengthExtended = emptyT<WORD>();
		std::shared_ptr<blockStringW> m_FolderList1 = emptySW();
		std::shared_ptr<blockT<DWORD>> m_FolderList2Length = emptyT<DWORD>();
		std::shared_ptr<EntryList> m_FolderList2;
		std::shared_ptr<blockT<DWORD>> m_AddressCount = emptyT<DWORD>(); // SFST_BINARY
		std::vector<std::shared_ptr<AddressListEntryStruct>> m_Addresses; // SFST_BINARY
		std::shared_ptr<blockT<DWORD>> m_SkipLen2 = emptyT<DWORD>();
		std::shared_ptr<blockBytes> m_SkipBytes2 = emptyBB();
		std::shared_ptr<RestrictionStruct> m_Restriction; // SFST_MRES
		std::shared_ptr<blockBytes> m_AdvancedSearchBytes = emptyBB(); // SFST_FILTERSTREAM
		std::shared_ptr<blockT<DWORD>> m_SkipLen3 = emptyT<DWORD>();
		std::shared_ptr<blockBytes> m_SkipBytes3 = emptyBB();
	};
} // namespace smartview