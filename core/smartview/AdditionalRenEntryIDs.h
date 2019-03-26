#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct PersistElement
	{
		static const WORD ELEMENT_SENTINEL = 0;
		std::shared_ptr<blockT<WORD>> wElementID = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wElementDataSize = emptyT<WORD>();
		std::shared_ptr<blockBytes> lpbElementData = emptyBB();

		PersistElement(const std::shared_ptr<binaryParser>& parser);
	};

	struct PersistData
	{
		static const WORD PERISIST_SENTINEL = 0;
		std::shared_ptr<blockT<WORD>> wPersistID = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> wDataElementsSize = emptyT<WORD>();
		std::vector<std::shared_ptr<PersistElement>> ppeDataElement;

		std::shared_ptr<blockBytes> JunkData = emptyBB();

		PersistData(const std::shared_ptr<binaryParser>& parser);
	};

	class AdditionalRenEntryIDs : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::vector<std::shared_ptr<PersistData>> m_ppdPersistData;
	};
} // namespace smartview