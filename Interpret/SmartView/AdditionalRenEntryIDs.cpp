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
			const auto wPersistID = m_Parser.Get<WORD>();
			const auto wDataElementSize = m_Parser.Get<WORD>();
			// Must have at least wDataElementSize bytes left to be a valid data element
			if (m_Parser.RemainingBytes() < wDataElementSize) break;

			m_Parser.advance(wDataElementSize);
			wPersistDataCount++;
			if (wPersistID == PERISIST_SENTINEL) break;
		}

		// Now we parse for real
		m_Parser.rewind();

		if (wPersistDataCount && wPersistDataCount < _MaxEntriesSmall)
		{
			m_ppdPersistData.reserve(wPersistDataCount);
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

		if (persistData.wPersistID != PERISIST_SENTINEL && m_Parser.RemainingBytes() >= persistData.wDataElementsSize)
		{
			// Build a new m_Parser to preread and count our elements
			// This new m_Parser will only contain as much space as suggested in wDataElementsSize
			CBinaryParser DataElementParser(persistData.wDataElementsSize, m_Parser.GetCurrentAddress());
			for (;;)
			{
				if (DataElementParser.RemainingBytes() < 2 * sizeof(WORD)) break;
				const auto wElementID = DataElementParser.Get<WORD>();
				const auto wElementDataSize = DataElementParser.Get<WORD>();
				// Must have at least wElementDataSize bytes left to be a valid element data
				if (DataElementParser.RemainingBytes() < wElementDataSize) break;

				DataElementParser.advance(wElementDataSize);
				wDataElementCount++;
				if (wElementID == ELEMENT_SENTINEL) break;
			}
		}

		if (wDataElementCount && wDataElementCount < _MaxEntriesSmall)
		{
			persistData.ppeDataElement.reserve(wDataElementCount);
			for (WORD iDataElement = 0; iDataElement < wDataElementCount; iDataElement++)
			{
				PersistElement persistElement;
				persistElement.wElementID = m_Parser.Get<WORD>();
				persistElement.wElementDataSize = m_Parser.Get<WORD>();
				if (persistElement.wElementID == ELEMENT_SENTINEL) break;
				// Since this is a word, the size will never be too large
				persistElement.lpbElementData = m_Parser.GetBYTES(persistElement.wElementDataSize);

				persistData.ppeDataElement.push_back(persistElement);
			}
		}

		// We'll trust wDataElementsSize to dictate our record size.
		// Count the 2 WORD size header fields too.
		const auto cbRecordSize = persistData.wDataElementsSize + sizeof(WORD) * 2;

		// Junk data remains - can't use GetRemainingData here since it would eat the whole buffer
		if (m_Parser.GetCurrentOffset() < cbRecordSize)
		{
			persistData.JunkData = m_Parser.GetBYTES(cbRecordSize - m_Parser.GetCurrentOffset());
		}

		return persistData;
	}

	void AdditionalRenEntryIDs::ParseBlocks()
	{
		setRoot(L"Additional Ren Entry IDs\r\n");
		addHeader(L"PersistDataCount = %1!d!", m_ppdPersistData.size());

		if (m_ppdPersistData.size())
		{
			for (WORD iPersistElement = 0; iPersistElement < m_ppdPersistData.size(); iPersistElement++)
			{
				terminateBlock();
				addBlankLine();
				auto element = block{};
				element.setText(L"Persist Element %1!d!:\r\n", iPersistElement);
				element.addBlock(
					m_ppdPersistData[iPersistElement].wPersistID,
					L"PersistID = 0x%1!04X! = %2!ws!\r\n",
					m_ppdPersistData[iPersistElement].wPersistID.getData(),
					interpretprop::InterpretFlags(flagPersistID, m_ppdPersistData[iPersistElement].wPersistID).c_str());
				element.addBlock(
					m_ppdPersistData[iPersistElement].wDataElementsSize,
					L"DataElementsSize = 0x%1!04X!",
					m_ppdPersistData[iPersistElement].wDataElementsSize.getData());

				if (m_ppdPersistData[iPersistElement].ppeDataElement.size())
				{
					for (WORD iDataElement = 0; iDataElement < m_ppdPersistData[iPersistElement].ppeDataElement.size();
						 iDataElement++)
					{
						element.terminateBlock();
						element.addHeader(L"DataElement: %1!d!\r\n", iDataElement);

						element.addBlock(
							m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID,
							L"\tElementID = 0x%1!04X! = %2!ws!\r\n",
							m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID.getData(),
							interpretprop::InterpretFlags(
								flagElementID,
								m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementID)
								.c_str());

						element.addBlock(
							m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementDataSize,
							L"\tElementDataSize = 0x%1!04X!\r\n",
							m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].wElementDataSize.getData());

						element.addHeader(L"\tElementData = ");
						element.addBlock(m_ppdPersistData[iPersistElement].ppeDataElement[iDataElement].lpbElementData);
					}
				}

				if (!m_ppdPersistData[iPersistElement].JunkData.empty())
				{
					element.terminateBlock();
					element.addHeader(
						L"Unparsed data size = 0x%1!08X!\r\n", m_ppdPersistData[iPersistElement].JunkData.size());
					element.addBlock(m_ppdPersistData[iPersistElement].JunkData);
				}

				addBlock(element);
			}
		}
	}
} // namespace smartview