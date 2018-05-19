#pragma once
#include "SmartViewParser.h"

namespace smartview
{
	struct PersistElement
	{
		WORD wElementID;
		WORD wElementDataSize;
		std::vector<BYTE> lpbElementData;
	};

	struct PersistData
	{
		WORD wPersistID;
		WORD wDataElementsSize;
		std::vector<PersistElement> ppeDataElement;

		std::vector<BYTE> JunkData;
	};

	class AdditionalRenEntryIDs : public SmartViewParser
	{
	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;
		PersistData BinToPersistData();

		std::vector<PersistData> m_ppdPersistData;
	};
}