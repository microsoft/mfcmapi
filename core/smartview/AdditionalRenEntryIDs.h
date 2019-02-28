#pragma once
#include <core/smartview/SmartViewParser.h>

namespace smartview
{
	struct PersistElement
	{
		blockT<WORD> wElementID;
		blockT<WORD> wElementDataSize;
		blockBytes lpbElementData;
	};

	struct PersistData
	{
		blockT<WORD> wPersistID;
		blockT<WORD> wDataElementsSize;
		std::vector<PersistElement> ppeDataElement;

		blockBytes JunkData;
	};

	class AdditionalRenEntryIDs : public SmartViewParser
	{
	private:
		void Parse() override;
		PersistData BinToPersistData();
		void ParseBlocks() override;

		std::vector<PersistData> m_ppdPersistData;
	};
} // namespace smartview