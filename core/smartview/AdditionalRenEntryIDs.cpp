#include <core/stdafx.h>
#include <core/smartview/AdditionalRenEntryIDs.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void AdditionalRenEntryIDs::Parse()
	{
		WORD wPersistDataCount = 0;
		// Run through the parser once to count the number of PersistData structs
		while (m_Parser->RemainingBytes() >= 2 * sizeof(WORD))
		{
			const auto wPersistID = m_Parser->Get<WORD>();
			const auto wDataElementSize = m_Parser->Get<WORD>();
			// Must have at least wDataElementSize bytes left to be a valid data element
			if (m_Parser->RemainingBytes() < wDataElementSize) break;

			m_Parser->advance(wDataElementSize);
			wPersistDataCount++;
			if (wPersistID == PERISIST_SENTINEL) break;
		}

		// Now we parse for real
		m_Parser->rewind();

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
		persistData.wPersistID = m_Parser->Get<WORD>();
		persistData.wDataElementsSize = m_Parser->Get<WORD>();

		if (persistData.wPersistID != PERISIST_SENTINEL && m_Parser->RemainingBytes() >= persistData.wDataElementsSize)
		{
			// Build a new m_Parser to preread and count our elements
			// This new m_Parser will only contain as much space as suggested in wDataElementsSize
			binaryParser DataElementParser(persistData.wDataElementsSize, m_Parser->GetCurrentAddress());
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
				persistElement.wElementID = m_Parser->Get<WORD>();
				persistElement.wElementDataSize = m_Parser->Get<WORD>();
				if (persistElement.wElementID == ELEMENT_SENTINEL) break;
				// Since this is a word, the size will never be too large
				persistElement.lpbElementData = m_Parser->GetBYTES(persistElement.wElementDataSize);

				persistData.ppeDataElement.push_back(persistElement);
			}
		}

		// We'll trust wDataElementsSize to dictate our record size.
		// Count the 2 WORD size header fields too.
		const auto cbRecordSize = persistData.wDataElementsSize + sizeof(WORD) * 2;

		// Junk data remains - can't use GetRemainingData here since it would eat the whole buffer
		if (m_Parser->GetCurrentOffset() < cbRecordSize)
		{
			persistData.JunkData = m_Parser->GetBYTES(cbRecordSize - m_Parser->GetCurrentOffset());
		}

		return persistData;
	}

	void AdditionalRenEntryIDs::ParseBlocks()
	{
		setRoot(L"Additional Ren Entry IDs\r\n");
		addHeader(L"PersistDataCount = %1!d!", m_ppdPersistData.size());

		if (!m_ppdPersistData.empty())
		{
			auto iPersistElement = WORD{};
			for (const auto& data : m_ppdPersistData)
			{
				terminateBlock();
				addBlankLine();
				auto element = block{};
				element.setText(L"Persist Element %1!d!:\r\n", iPersistElement);
				element.addBlock(
					data.wPersistID,
					L"PersistID = 0x%1!04X! = %2!ws!\r\n",
					data.wPersistID.getData(),
					flags::InterpretFlags(flagPersistID, data.wPersistID).c_str());
				element.addBlock(
					data.wDataElementsSize, L"DataElementsSize = 0x%1!04X!", data.wDataElementsSize.getData());

				if (!data.ppeDataElement.empty())
				{
					auto iDataElement = WORD{};
					for (auto& dataElement : data.ppeDataElement)
					{
						element.terminateBlock();
						element.addHeader(L"DataElement: %1!d!\r\n", iDataElement);

						element.addBlock(
							dataElement.wElementID,
							L"\tElementID = 0x%1!04X! = %2!ws!\r\n",
							dataElement.wElementID.getData(),
							flags::InterpretFlags(flagElementID, dataElement.wElementID).c_str());

						element.addBlock(
							dataElement.wElementDataSize,
							L"\tElementDataSize = 0x%1!04X!\r\n",
							dataElement.wElementDataSize.getData());

						element.addHeader(L"\tElementData = ");
						element.addBlock(dataElement.lpbElementData);
						iDataElement++;
					}
				}

				if (!data.JunkData.empty())
				{
					element.terminateBlock();
					element.addHeader(L"Unparsed data size = 0x%1!08X!\r\n", data.JunkData.size());
					element.addBlock(data.JunkData);
				}

				addBlock(element);
				iPersistElement++;
			}
		}
	}
} // namespace smartview