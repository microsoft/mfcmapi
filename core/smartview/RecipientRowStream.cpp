#include <core/stdafx.h>
#include <core/smartview/RecipientRowStream.h>
#include <core/smartview/PropertiesStruct.h>

namespace smartview
{
	ADRENTRYStruct::ADRENTRYStruct(const std::shared_ptr<binaryParser>& parser)
	{
		cValues = blockT<DWORD>::parse(parser);
		ulReserved1 = blockT<DWORD>::parse(parser);

		if (*cValues)
		{
			if (*cValues < _MaxEntriesSmall)
			{
				rgPropVals.SetMaxEntries(*cValues);
				rgPropVals.SmartViewParser::parse(parser, false);
			}
		}
	}

	void RecipientRowStream::Parse()
	{
		m_cVersion = blockT<DWORD>::parse(m_Parser);
		m_cRowCount = blockT<DWORD>::parse(m_Parser);

		if (*m_cRowCount)
		{
			if (*m_cRowCount < _MaxEntriesSmall)
			{
				m_lpAdrEntry.reserve(*m_cRowCount);
				for (DWORD i = 0; i < *m_cRowCount; i++)
				{
					m_lpAdrEntry.emplace_back(std::make_shared<ADRENTRYStruct>(m_Parser));
				}
			}
		}
	}

	void RecipientRowStream::ParseBlocks()
	{
		setRoot(L"Recipient Row Stream\r\n");
		addChild(m_cVersion, L"cVersion = %1!d!\r\n", m_cVersion->getData());
		addChild(m_cRowCount, L"cRowCount = %1!d!\r\n", m_cRowCount->getData());
		if (!m_lpAdrEntry.empty())
		{
			addBlankLine();
			auto i = DWORD{};
			for (const auto& entry : m_lpAdrEntry)
			{
				terminateBlock();
				addHeader(L"Row %1!d!\r\n", i);
				addChild(entry->cValues, L"cValues = 0x%1!08X! = %1!d!\r\n", entry->cValues->getData());
				addChild(entry->ulReserved1, L"ulReserved1 = 0x%1!08X! = %1!d!\r\n", entry->ulReserved1->getData());
				addChild(entry->rgPropVals.getBlock());

				i++;
			}
		}
	}
} // namespace smartview