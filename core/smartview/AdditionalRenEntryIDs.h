#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class PersistElement : public block
	{
	public:
		static const WORD ELEMENT_SENTINEL = 0;

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> wElementID = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wElementDataSize = emptyT<WORD>();
		std::shared_ptr<blockBytes> lpbElementData = emptyBB();
	};

	class PersistData : public block
	{
	public:
		static const WORD PERISIST_SENTINEL = 0;

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> wPersistID = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wDataElementsSize = emptyT<WORD>();
		std::vector<std::shared_ptr<PersistElement>> ppeDataElement;

		std::shared_ptr<blockBytes> junkData = emptyBB();
	};

	class AdditionalRenEntryIDs : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::vector<std::shared_ptr<PersistData>> m_ppdPersistData;
	};
} // namespace smartview