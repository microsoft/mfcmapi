#include "stdafx.h"
#include "AdditionalRenEntryIDs.h"
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>

#define PERISIST_SENTINEL 0
#define ELEMENT_SENTINEL 0

void AdditionalRenEntryIDs::Parse()
{
	WORD wPersistDataCount = 0;
	// Run through the parser once to count the number of PersistData structs
	for (;;)
	{
		if (m_Parser.RemainingBytes() < 2 * sizeof(WORD)) break;
		auto wPersistID = m_Parser.Get<WORD>();
		auto wDataElementSize = m_Parser.Get<WORD>();
		// Must have at least wDataElementSize bytes left to be a valid data element
		if (m_Parser.RemainingBytes() < wDataElementSize) break;

		m_Parser.Advance(wDataElementSize);
		wPersistDataCount++;
		if (PERISIST_SENTINEL == wPersistID) break;
	}

	// Now we parse for real
	m_Parser.Rewind();

	if (wPersistDataCount && wPersistDataCount < _MaxEntriesSmall)
	{
		for (WORD iPersistElement = 0; iPersistElement < wPersistDataCount; iPersistElement++)
		{
			m_ppdPersistData.push_back(BinToPersistData());
		}
	}
}

PersistData AdditionalRenEntryIDs::BinToPersistData()
{
	PersistData persistData;
	WORD wDataElementCount = 0;
	persistData.wPersistID = m_Parser.Get<WORD>();
	persistData.wDataElementsSize = m_Parser.Get<WORD>();

	if (persistData.wPersistID != PERISIST_SENTINEL &&
		m_Parser.RemainingBytes() >= persistData.wDataElementsSize)
	{
		// Build a new m_Parser to preread and count our elements
		// This new m_Parser will only contain as much space as suggested in wDataElementsSize
		CBinaryParser DataElementParser(persistData.wDataElementsSize, m_Parser.GetCurrentAddress());
		for (;;)
		{
			if (DataElementParser.RemainingBytes() < 2 * sizeof(WORD)) break;
			auto wElementID = DataElementParser.Get<WORD>();
			auto wElementDataSize = DataElementParser.Get<WORD>();
			// Must have at least wElementDataSize bytes left to be a valid element data
			if (DataElementParser.RemainingBytes() < wElementDataSize) break;

			DataElementParser.Advance(wElementDataSize);
			wDataElementCount++;
			if (ELEMENT_SENTINEL == wElementID) break;
		}
	}

	if (wDataElementCount && wDataElementCount < _MaxEntriesSmall)
	{
		for (WORD iDataElement = 0; iDataElement < wDataElementCount; iDataElement++)
		{
			PersistElement persistElement;
			persistElement.wElementID = m_Parser.Get<WORD>();
			persistElement.wElementDataSize = m_Parser.Get<WORD>();
			if (ELEMENT_SENTINEL == persistElement.wElementID) break;
			// Since this is a word, the size will never be too large
			persistElement.lpbElementData = m_Parser.GetBYTES(persistElement.wElementDataSize);

			persistData.ppeDataElement.push_back(persistElement);
		}
	}

	// We'll trust wDataElementsSize to dictate our record size.
	// Count the 2 WORD size header fields too.
	auto cbRecordSize = persistData.wDataElementsSize + sizeof(WORD) * 2;

	// Junk data remains - can't use GetRemainingData here since it would eat the whole buffer
	if (m_Parser.GetCurrentOffset() < cbRecordSize)
	{
		persistData.JunkData = m_Parser.GetBYTES(cbRecordSize - m_Parser.GetCurrentOffset());
	}

	return persistData;
}

_Check_return_ std::wstring AdditionalRenEntryIDs::ToStringInternal()
{
	auto szAdditionalRenEntryIDs = strings::formatmessage(IDS_AEIDHEADER, m_ppdPersistData.size());

	if (m_ppdPersistData.size())
	{
		for (WORD iPersistElement = 0; iPersistElement < m_ppdPersistData.size(); iPersistElement++)
		{
			szAdditionalRenEntryIDs += strings::formatmessage(IDS_AEIDPERSISTELEMENT,
				iPersistElement,
				m_ppdPersistData[iPersistElement].wPersistID,
				InterpretFlags(flagPersistID, m_ppdPersistData[iPersistElement].wPersistID).c_str(),
				m_ppdPersistData[iPersistElement].wDataElementsSize);

			if (m_ppdPersistData[iPersistElement].ppeDataElement.size())
			{
				for (WORD iDataElement = 0; iDataElement < m_ppdPersistData[iPersistElement].ppeDataElement.size(); iDataElement++)
				{
					szAdditionalRenEntryIDs += strings::formatmessage(IDS_AEIDDATAELEMENT,
						iDataElement,
						m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID,
						InterpretFlags(flagElementID, m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID).c_str(),
						m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementDataSize);

					szAdditionalRenEntryIDs += strings::BinToHexString(m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].lpbElementData, true);
				}
			}

			szAdditionalRenEntryIDs += JunkDataToString(m_ppdPersistData[iPersistElement].JunkData);
		}
	}

	return szAdditionalRenEntryIDs;
}