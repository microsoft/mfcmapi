#include <core/stdafx.h>
#include <core/smartview/AdditionalRenEntryIDs.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	PersistElement::PersistElement(const std::shared_ptr<binaryParser>& parser)
	{
		wElementID = blockT<WORD>::parse(parser);
		wElementDataSize = blockT<WORD>::parse(parser);
		if (wElementID != PersistElement::ELEMENT_SENTINEL)
		{
			// Since this is a word, the size will never be too large
			lpbElementData = blockBytes::parse(parser, wElementDataSize->getData());
		}
	}

	void AdditionalRenEntryIDs::parse()
	{
		WORD wPersistDataCount = 0;
		// Run through the parser once to count the number of PersistData structs
		while (parser->getSize() >= 2 * sizeof(WORD))
		{
			const auto& wPersistID = blockT<WORD>::parse(parser);
			const auto& wDataElementSize = blockT<WORD>::parse(parser);
			// Must have at least wDataElementSize bytes left to be a valid data element
			if (parser->getSize() < *wDataElementSize) break;

			parser->advance(*wDataElementSize);
			wPersistDataCount++;
			if (wPersistID == PersistData::PERISIST_SENTINEL) break;
		}

		// Now we parse for real
		parser->rewind();

		if (wPersistDataCount && wPersistDataCount < _MaxEntriesSmall)
		{
			m_ppdPersistData.reserve(wPersistDataCount);
			for (WORD iPersistElement = 0; iPersistElement < wPersistDataCount; iPersistElement++)
			{
				m_ppdPersistData.emplace_back(std::make_shared<PersistData>(parser));
			}
		}
	}

	PersistData::PersistData(const std::shared_ptr<binaryParser>& parser)
	{
		WORD wDataElementCount = 0;
		wPersistID = blockT<WORD>::parse(parser);
		wDataElementsSize = blockT<WORD>::parse(parser);

		if (wPersistID != PERISIST_SENTINEL && parser->getSize() >= *wDataElementsSize)
		{
			// Build a new parser to preread and count our elements
			// This new parser will only contain as much space as suggested in wDataElementsSize
			auto DataElementParser = std::make_shared<binaryParser>(*wDataElementsSize, parser->getAddress());
			for (;;)
			{
				if (DataElementParser->getSize() < 2 * sizeof(WORD)) break;
				const auto& wElementID = blockT<WORD>::parse(DataElementParser);
				const auto& wElementDataSize = blockT<WORD>::parse(DataElementParser);
				// Must have at least wElementDataSize bytes left to be a valid element data
				if (DataElementParser->getSize() < *wElementDataSize) break;

				DataElementParser->advance(*wElementDataSize);
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
		const auto cbRecordSize = *wDataElementsSize + sizeof(WORD) * 2;

		// Junk data remains - can't use GetRemainingData here since it would eat the whole buffer
		if (parser->getOffset() < cbRecordSize)
		{
			JunkData = blockBytes::parse(parser, cbRecordSize - parser->getOffset());
		}
	}

	void AdditionalRenEntryIDs::parseBlocks()
	{
		setText(L"Additional Ren Entry IDs");
		addHeader(L"PersistDataCount = %1!d!", m_ppdPersistData.size());

		if (!m_ppdPersistData.empty())
		{
			auto iPersistElement = 0;
			for (const auto& persistData : m_ppdPersistData)
			{
				auto element = create(L"Persist Element %1!d!:", iPersistElement);
				addChild(element);

				element->addChild(
					persistData->wPersistID,
					L"PersistID = 0x%1!04X! = %2!ws!",
					persistData->wPersistID->getData(),
					flags::InterpretFlags(flagPersistID, persistData->wPersistID->getData()).c_str());
				element->addChild(
					persistData->wDataElementsSize,
					L"DataElementsSize = 0x%1!04X!",
					persistData->wDataElementsSize->getData());

				if (!persistData->ppeDataElement.empty())
				{
					auto iDataElement = 0;
					for (const auto& dataElement : persistData->ppeDataElement)
					{
						auto de = create(L"DataElement: %1!d!", iDataElement);
						element->addChild(de);

						de->addChild(
							dataElement->wElementID,
							L"ElementID = 0x%1!04X! = %2!ws!",
							dataElement->wElementID->getData(),
							flags::InterpretFlags(flagElementID, dataElement->wElementID->getData()).c_str());

						de->addChild(
							dataElement->wElementDataSize,
							L"ElementDataSize = 0x%1!04X!",
							dataElement->wElementDataSize->getData());

						de->addLabeledChild(L"ElementData = ", dataElement->lpbElementData);
						iDataElement++;
					}
				}

				if (!persistData->JunkData->empty())
				{
					element->addHeader(L"Unparsed data size = 0x%1!08X!", persistData->JunkData->size());
					element->addChild(persistData->JunkData);
				}

				iPersistElement++;
			}
		}
	}
} // namespace smartview