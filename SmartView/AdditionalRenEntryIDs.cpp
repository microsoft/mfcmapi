#include "stdafx.h"
#include "..\stdafx.h"
#include "AdditionalRenEntryIDs.h"
#include "..\String.h"
#include "..\InterpretProp2.h"
#include "..\ExtraPropTags.h"
#include "..\ParseProperty.h"

AdditionalRenEntryIDs::AdditionalRenEntryIDs(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_wPersistDataCount = 0;
	m_ppdPersistData = NULL;
}

AdditionalRenEntryIDs::~AdditionalRenEntryIDs()
{
	if (m_ppdPersistData)
	{
		WORD iPersistElement = 0;
		for (iPersistElement = 0; iPersistElement < m_wPersistDataCount; iPersistElement++)
		{
			if (m_ppdPersistData[iPersistElement].ppeDataElement)
			{
				WORD iDataElement = 0;
				for (iDataElement = 0; iDataElement < m_ppdPersistData[iPersistElement].wDataElementCount; iDataElement++)
				{
					delete[] m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].lpbElementData;
				}
			}

			delete[] m_ppdPersistData[iPersistElement].ppeDataElement;
			delete[] m_ppdPersistData[iPersistElement].JunkData;
		}
	}

	delete[] m_ppdPersistData;
}

#define PERISIST_SENTINEL 0
#define ELEMENT_SENTINEL 0

void AdditionalRenEntryIDs::Parse()
{
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
		m_wPersistDataCount++;
		if (PERISIST_SENTINEL == wPersistID) break;
	}

	// Now we parse for real
	m_Parser.Rewind();

	if (m_wPersistDataCount && m_wPersistDataCount < _MaxEntriesSmall)
	{
		m_ppdPersistData = new PersistData[m_wPersistDataCount];

		if (m_ppdPersistData)
		{
			memset(m_ppdPersistData, 0, m_wPersistDataCount * sizeof(PersistData));
			WORD iPersistElement = 0;
			for (iPersistElement = 0; iPersistElement < m_wPersistDataCount; iPersistElement++)
			{
				BinToPersistData(
					&m_ppdPersistData[iPersistElement]);
			}
		}
	}
}

void AdditionalRenEntryIDs::BinToPersistData(_Out_ PersistData* ppdPersistData)
{
	if (!ppdPersistData) return;

	m_Parser.GetWORD(&ppdPersistData->wPersistID);
	m_Parser.GetWORD(&ppdPersistData->wDataElementsSize);

	if (ppdPersistData->wPersistID != PERISIST_SENTINEL &&
		m_Parser.RemainingBytes() >= ppdPersistData->wDataElementsSize)
	{
		// Build a new m_Parser to preread and count our elements
		// This new m_Parser will only contain as much space as suggested in wDataElementsSize
		CBinaryParser DataElementParser(ppdPersistData->wDataElementsSize, m_Parser.GetCurrentAddress());
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
			ppdPersistData->wDataElementCount++;
			if (ELEMENT_SENTINEL == wElementID) break;
		}
	}

	if (ppdPersistData->wDataElementCount && ppdPersistData->wDataElementCount < _MaxEntriesSmall)
	{
		ppdPersistData->ppeDataElement = new PersistElement[ppdPersistData->wDataElementCount];

		if (ppdPersistData->ppeDataElement)
		{
			memset(ppdPersistData->ppeDataElement, 0, ppdPersistData->wDataElementCount * sizeof(PersistElement));

			WORD iDataElement = 0;
			for (iDataElement = 0; iDataElement < ppdPersistData->wDataElementCount; iDataElement++)
			{
				m_Parser.GetWORD(&ppdPersistData->ppeDataElement[iDataElement].wElementID);
				m_Parser.GetWORD(&ppdPersistData->ppeDataElement[iDataElement].wElementDataSize);
				if (ELEMENT_SENTINEL == ppdPersistData->ppeDataElement[iDataElement].wElementID) break;
				// Since this is a word, the size will never be too large
				m_Parser.GetBYTES(
					ppdPersistData->ppeDataElement[iDataElement].wElementDataSize,
					ppdPersistData->ppeDataElement[iDataElement].wElementDataSize,
					&ppdPersistData->ppeDataElement[iDataElement].lpbElementData);
			}
		}
	}

	// We'll trust wDataElementsSize to dictate our record size.
	// Count the 2 WORD size header fields too.
	size_t cbRecordSize = ppdPersistData->wDataElementsSize + sizeof(WORD)* 2;

	// Junk data remains - can't use GetRemainingData here since it would eat the whole buffer
	if (m_Parser.GetCurrentOffset() < cbRecordSize)
	{
		ppdPersistData->JunkDataSize = cbRecordSize - m_Parser.GetCurrentOffset();
		m_Parser.GetBYTES(ppdPersistData->JunkDataSize, ppdPersistData->JunkDataSize, &ppdPersistData->JunkData);
	}
}

_Check_return_ wstring AdditionalRenEntryIDs::ToStringInternal()
{
	wstring szAdditionalRenEntryIDs;
	wstring szTmp;

	szAdditionalRenEntryIDs = formatmessage(IDS_AEIDHEADER, m_wPersistDataCount);

	if (m_ppdPersistData)
	{
		WORD iPersistElement = 0;
		for (iPersistElement = 0; iPersistElement < m_wPersistDataCount; iPersistElement++)
		{
			wstring szPersistID = InterpretFlags(flagPersistID, m_ppdPersistData[iPersistElement].wPersistID);
			szTmp = formatmessage(IDS_AEIDPERSISTELEMENT,
				iPersistElement,
				m_ppdPersistData[iPersistElement].wPersistID, szPersistID.c_str(),
				m_ppdPersistData[iPersistElement].wDataElementsSize);
			szAdditionalRenEntryIDs += szTmp;

			if (m_ppdPersistData[iPersistElement].ppeDataElement)
			{
				WORD iDataElement = 0;
				for (iDataElement = 0; iDataElement < m_ppdPersistData[iPersistElement].wDataElementCount; iDataElement++)
				{
					wstring szElementID = InterpretFlags(flagElementID, m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID);
					szTmp = formatmessage(IDS_AEIDDATAELEMENT,
						iDataElement,
						m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID, szElementID.c_str(),
						m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementDataSize);
					szAdditionalRenEntryIDs += szTmp;

					SBinary sBin = { 0 };
					sBin.cb = (ULONG)m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementDataSize;
					sBin.lpb = m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].lpbElementData;
					szAdditionalRenEntryIDs += BinToHexString(&sBin, true);
				}
			}

			szAdditionalRenEntryIDs += JunkDataToString(m_ppdPersistData[iPersistElement].JunkDataSize, m_ppdPersistData[iPersistElement].JunkData);
		}
	}

	return szAdditionalRenEntryIDs;
}