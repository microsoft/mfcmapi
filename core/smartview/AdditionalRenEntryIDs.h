#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	struct PersistElement
	{
		static const WORD ELEMENT_SENTINEL = 0;
		blockT<WORD> wElementID;
		blockT<WORD> wElementDataSize;
		blockBytes lpbElementData;

		PersistElement(std::shared_ptr<binaryParser> parser);
	};

	struct PersistData
	{
		static const WORD PERISIST_SENTINEL = 0;
		blockT<WORD> wPersistID;
		blockT<WORD> wDataElementsSize;
		std::vector<std::shared_ptr<PersistElement>> ppeDataElement;

		blockBytes JunkData;

		PersistData(std::shared_ptr<binaryParser> parser);
	};

	class AdditionalRenEntryIDs : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		std::vector<std::shared_ptr<PersistData>> m_ppdPersistData;
	};
} // namespace smartview