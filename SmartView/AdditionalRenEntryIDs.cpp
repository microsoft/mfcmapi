#include "stdafx.h"
#include "AdditionalRenEntryIDs.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

AdditionalRenEntryIDs::AdditionalRenEntryIDs(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin) {}

#define PERISIST_SENTINEL 0
#define ELEMENT_SENTINEL 0

void AdditionalRenEntryIDs::Parse()
{
	WORD wPersistDataCount = 0;
	// Run through the parser once to count the number of PersistData structs
	for (;;)
	{
		if (m_Parser.RemainingBytes() < 2 * sizeof(WORD)) break;
		WORD wPersistID = NULL;
		WORD wDataElementSize = NULL;
		m_Parser.GetWORD(&wPersistID);
		m_Parser.GetWORD(&wDataElementSize);
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
	m_Parser.GetWORD(&persistData.wPersistID);
	m_Parser.GetWORD(&persistData.wDataElementsSize);

	if (persistData.wPersistID != PERISIST_SENTINEL &&
		m_Parser.RemainingBytes() >= persistData.wDataElementsSize)
	{
		// Build a new m_Parser to preread and count our elements
		// This new m_Parser will only contain as much space as suggested in wDataElementsSize
		CBinaryParser DataElementParser(persistData.wDataElementsSize, m_Parser.GetCurrentAddress());
		for (;;)
		{
			if (DataElementParser.RemainingBytes() < 2 * sizeof(WORD)) break;
			WORD wElementID = NULL;
			WORD wElementDataSize = NULL;
			DataElementParser.GetWORD(&wElementID);
			DataElementParser.GetWORD(&wElementDataSize);
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
			m_Parser.GetWORD(&persistElement.wElementID);
			m_Parser.GetWORD(&persistElement.wElementDataSize);
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

_Check_return_ wstring AdditionalRenEntryIDs::ToStringInternal()
{
	auto szAdditionalRenEntryIDs = formatmessage(IDS_AEIDHEADER, m_ppdPersistData.size());

	if (m_ppdPersistData.size())
	{
		for (WORD iPersistElement = 0; iPersistElement < m_ppdPersistData.size(); iPersistElement++)
		{
			szAdditionalRenEntryIDs += formatmessage(IDS_AEIDPERSISTELEMENT,
				iPersistElement,
				m_ppdPersistData[iPersistElement].wPersistID,
				InterpretFlags(flagPersistID, m_ppdPersistData[iPersistElement].wPersistID).c_str(),
				m_ppdPersistData[iPersistElement].wDataElementsSize);

			if (m_ppdPersistData[iPersistElement].ppeDataElement.size())
			{
				for (WORD iDataElement = 0; iDataElement < m_ppdPersistData[iPersistElement].ppeDataElement.size(); iDataElement++)
				{
					szAdditionalRenEntryIDs += formatmessage(IDS_AEIDDATAELEMENT,
						iDataElement,
						m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID,
						InterpretFlags(flagElementID, m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID).c_str(),
						m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementDataSize);

					szAdditionalRenEntryIDs += BinToHexString(m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].lpbElementData, true);
				}
			}

			szAdditionalRenEntryIDs += JunkDataToString(m_ppdPersistData[iPersistElement].JunkData);
		}
	}

	return szAdditionalRenEntryIDs;
}