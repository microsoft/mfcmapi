#include <StdAfx.h>
#include <Interpret/SmartView/AdditionalRenEntryIDs.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
#define PERISIST_SENTINEL 0
#define ELEMENT_SENTINEL 0

	void AdditionalRenEntryIDs::Parse()
	{
		WORD wPersistDataCount = 0;
		// Run through the parser once to count the number of PersistData structs
		for (;;)
		{
			if (m_Parser.RemainingBytes() < 2 * sizeof(WORD)) break;
			const auto wPersistID = m_Parser.GetBlock<WORD>();
			const auto wDataElementSize = m_Parser.GetBlock<WORD>();
			// Must have at least wDataElementSize bytes left to be a valid data element
			if (m_Parser.RemainingBytes() < wDataElementSize.getData()) break;

			m_Parser.Advance(wDataElementSize.getData());
			wPersistDataCount++;
			if (wPersistID.getData() == PERISIST_SENTINEL) break;
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

		ParseBlocks();
	}

	PersistData AdditionalRenEntryIDs::BinToPersistData()
	{
		PersistData persistData;
		WORD wDataElementCount = 0;
		persistData.wPersistID = m_Parser.GetBlock<WORD>();
		persistData.wDataElementsSize = m_Parser.GetBlock<WORD>();

		if (persistData.wPersistID.getData() != PERISIST_SENTINEL &&
			m_Parser.RemainingBytes() >= persistData.wDataElementsSize.getData())
		{
			// Build a new m_Parser to preread and count our elements
			// This new m_Parser will only contain as much space as suggested in wDataElementsSize
			CBinaryParser DataElementParser(persistData.wDataElementsSize.getData(), m_Parser.GetCurrentAddress());
			for (;;)
			{
				if (DataElementParser.RemainingBytes() < 2 * sizeof(WORD)) break;
				const auto wElementID = DataElementParser.GetBlock<WORD>();
				const auto wElementDataSize = DataElementParser.GetBlock<WORD>();
				// Must have at least wElementDataSize bytes left to be a valid element data
				if (DataElementParser.RemainingBytes() < wElementDataSize.getData()) break;

				DataElementParser.Advance(wElementDataSize.getData());
				wDataElementCount++;
				if (wElementID.getData() == ELEMENT_SENTINEL) break;
			}
		}

		if (wDataElementCount && wDataElementCount < _MaxEntriesSmall)
		{
			for (WORD iDataElement = 0; iDataElement < wDataElementCount; iDataElement++)
			{
				PersistElement persistElement;
				persistElement.wElementID = m_Parser.GetBlock<WORD>();
				persistElement.wElementDataSize = m_Parser.GetBlock<WORD>();
				if (persistElement.wElementID.getData() == ELEMENT_SENTINEL) break;
				// Since this is a word, the size will never be too large
				persistElement.lpbElementData = m_Parser.GetBlockBYTES(persistElement.wElementDataSize.getData());

				persistData.ppeDataElement.push_back(persistElement);
			}
		}

		// We'll trust wDataElementsSize to dictate our record size.
		// Count the 2 WORD size header fields too.
		const auto cbRecordSize = persistData.wDataElementsSize.getData() + sizeof(WORD) * 2;

		// Junk data remains - can't use GetRemainingData here since it would eat the whole buffer
		if (m_Parser.GetCurrentOffset() < cbRecordSize)
		{
			persistData.JunkData = m_Parser.GetBlockBYTES(cbRecordSize - m_Parser.GetCurrentOffset());
		}

		return persistData;
	}

	void AdditionalRenEntryIDs::ParseBlocks()
	{
		addHeader(L"Additional Ren Entry IDs\r\n");
		addHeader(strings::formatmessage(L"PersistDataCount = %1!d!", m_ppdPersistData.size()));

		if (m_ppdPersistData.size())
		{
			for (WORD iPersistElement = 0; iPersistElement < m_ppdPersistData.size(); iPersistElement++)
			{
				addHeader(strings::formatmessage(L"\r\n\r\n"));
				addHeader(strings::formatmessage(L"Persist Element %1!d!:\r\n", iPersistElement));
				addBlock(
					m_ppdPersistData[iPersistElement].wPersistID,
					strings::formatmessage(
						L"PersistID = 0x%1!04X! = %2!ws!\r\n",
						m_ppdPersistData[iPersistElement].wPersistID.getData(),
						interpretprop::InterpretFlags(
							flagPersistID, m_ppdPersistData[iPersistElement].wPersistID.getData())
							.c_str()));
				addBlock(
					m_ppdPersistData[iPersistElement].wDataElementsSize,
					strings::formatmessage(
						L"DataElementsSize = 0x%1!04X!",
						m_ppdPersistData[iPersistElement].wDataElementsSize.getData()));

				if (m_ppdPersistData[iPersistElement].ppeDataElement.size())
				{
					for (WORD iDataElement = 0; iDataElement < m_ppdPersistData[iPersistElement].ppeDataElement.size();
						 iDataElement++)
					{
						addHeader(strings::formatmessage(L"\r\nDataElement: %1!d!\r\n", iDataElement));

						addBlock(
							m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID,
							strings::formatmessage(
								L"\tElementID = 0x%1!04X! = %2!ws!\r\n",
								m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID.getData(),
								interpretprop::InterpretFlags(
									flagElementID,
									m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID.getData())
									.c_str()));

						addBlock(
							m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementDataSize,
							strings::formatmessage(
								L"\tElementDataSize = 0x%1!04X!\r\n",
								m_ppdPersistData[iPersistElement]
									.ppeDataElement[iDataElement]
									.wElementDataSize.getData()));

						addHeader(strings::formatmessage(L"\tElementData = "));
						addBlock(
							m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].lpbElementData,
							strings::BinToHexString(
								m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].lpbElementData.getData(),
								true));
					}
				}

				addHeader(JunkDataToString(m_ppdPersistData[iPersistElement].JunkData.getData()));
			}
		}
	}
}