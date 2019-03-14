#include <core/stdafx.h>
#include <core/smartview/AdditionalRenEntryIDs.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	PersistElement::PersistElement(std::shared_ptr<binaryParser> parser)
	{
		wElementID.parse<WORD>(parser);
		wElementDataSize.parse<WORD>(parser);
		if (wElementID != PersistElement::ELEMENT_SENTINEL)
		{
			// Since this is a word, the size will never be too large
			lpbElementData.parse(parser, wElementDataSize);
		}
	}

	void AdditionalRenEntryIDs::Parse()
	{
		WORD wPersistDataCount = 0;
		// Run through the parser once to count the number of PersistData structs
		while (m_Parser->RemainingBytes() >= 2 * sizeof(WORD))
		{
			const auto& wPersistID = blockT<WORD>(m_Parser);
			const auto& wDataElementSize = blockT<WORD>(m_Parser);
			// Must have at least wDataElementSize bytes left to be a valid data element
			if (m_Parser->RemainingBytes() < wDataElementSize) break;

			m_Parser->advance(wDataElementSize);
			wPersistDataCount++;
			if (wPersistID == PersistData::PERISIST_SENTINEL) break;
		}

		// Now we parse for real
		m_Parser->rewind();

		if (wPersistDataCount && wPersistDataCount < _MaxEntriesSmall)
		{
			m_ppdPersistData.reserve(wPersistDataCount);
			for (WORD iPersistElement = 0; iPersistElement < wPersistDataCount; iPersistElement++)
			{
				m_ppdPersistData.emplace_back(std::make_shared<PersistData>(m_Parser));
			}
		}
	}

	PersistData::PersistData(std::shared_ptr<binaryParser> parser)
	{
		WORD wDataElementCount = 0;
		wPersistID.parse<WORD>(parser);
		wDataElementsSize.parse<WORD>(parser);

		if (wPersistID != PERISIST_SENTINEL && parser->RemainingBytes() >= wDataElementsSize)
		{
			// Build a new parser to preread and count our elements
			// This new parser will only contain as much space as suggested in wDataElementsSize
			auto DataElementParser = std::make_shared<binaryParser>(wDataElementsSize, parser->GetCurrentAddress());
			for (;;)
			{
				if (DataElementParser->RemainingBytes() < 2 * sizeof(WORD)) break;
				const auto& wElementID = blockT<WORD>(DataElementParser);
				const auto& wElementDataSize = blockT<WORD>(DataElementParser);
				// Must have at least wElementDataSize bytes left to be a valid element data
				if (DataElementParser->RemainingBytes() < wElementDataSize) break;

				DataElementParser->advance(wElementDataSize);
				wDataElementCount++;
				if (wElementID == PersistElement::ELEMENT_SENTINEL) break;
			}
		}

		if (wDataElementCount && wDataElementCount < _MaxEntriesSmall)
		{
			ppeDataElement.reserve(wDataElementCount);
			for (WORD iDataElement = 0; iDataElement < wDataElementCount; iDataElement++)
			{
				ppeDataElement.emplace_back(std::make_shared<PersistElement>(parser));
			}
		}

		// We'll trust wDataElementsSize to dictate our record size.
		// Count the 2 WORD size header fields too.
		const auto cbRecordSize = wDataElementsSize + sizeof(WORD) * 2;

		// Junk data remains - can't use GetRemainingData here since it would eat the whole buffer
		if (parser->GetCurrentOffset() < cbRecordSize)
		{
			JunkData.parse(parser, cbRecordSize - parser->GetCurrentOffset());
		}
	}

	void AdditionalRenEntryIDs::ParseBlocks()
	{
		setRoot(L"Additional Ren Entry IDs\r\n");
		addHeader(L"PersistDataCount = %1!d!", m_ppdPersistData.size());

		if (!m_ppdPersistData.empty())
		{
			auto iPersistElement = 0;
			for (const auto& persistData : m_ppdPersistData)
			{
				terminateBlock();
				addBlankLine();
				auto element = std::make_shared<block>();
				element->setText(L"Persist Element %1!d!:\r\n", iPersistElement);
				element->addChild(
					persistData->wPersistID,
					L"PersistID = 0x%1!04X! = %2!ws!\r\n",
					persistData->wPersistID.getData(),
					flags::InterpretFlags(flagPersistID, persistData->wPersistID).c_str());
				element->addChild(
					persistData->wDataElementsSize,
					L"DataElementsSize = 0x%1!04X!",
					persistData->wDataElementsSize.getData());

				if (!persistData->ppeDataElement.empty())
				{
					auto iDataElement = 0;
					for (const auto& dataElement : persistData->ppeDataElement)
					{
						element->terminateBlock();
						element->addHeader(L"DataElement: %1!d!\r\n", iDataElement);

						element->addChild(
							dataElement->wElementID,
							L"\tElementID = 0x%1!04X! = %2!ws!\r\n",
							dataElement->wElementID.getData(),
							flags::InterpretFlags(flagElementID, dataElement->wElementID).c_str());

						element->addChild(
							dataElement->wElementDataSize,
							L"\tElementDataSize = 0x%1!04X!\r\n",
							dataElement->wElementDataSize.getData());

						element->addHeader(L"\tElementData = ");
						element->addChild(dataElement->lpbElementData);
						iDataElement++;
					}
				}

				if (!persistData->JunkData.empty())
				{
					element->terminateBlock();
					element->addHeader(L"Unparsed data size = 0x%1!08X!\r\n", persistData->JunkData.size());
					element->addChild(persistData->JunkData);
				}

				addChild(element);
				iPersistElement++;
			}
		}
	}
} // namespace smartview