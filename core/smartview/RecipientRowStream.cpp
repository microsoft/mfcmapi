#include <core/stdafx.h>
#include <core/smartview/RecipientRowStream.h>

namespace smartview
{
	void ADRENTRYStruct::parse()
	{
		cValues = blockT<DWORD>::parse(parser);
		ulReserved1 = blockT<DWORD>::parse(parser);

		if (*cValues)
		{
			if (*cValues < _MaxEntriesSmall)
			{
				rgPropVals = std::make_shared<PropertiesStruct>(*cValues, false, false);
				rgPropVals->block::parse(parser, false);
			}
		}
	}

	void ADRENTRYStruct::parseBlocks()
	{
		addChild(cValues, L"cValues = 0x%1!08X! = %1!d!", cValues->getData());
		addChild(ulReserved1, L"ulReserved1 = 0x%1!08X! = %1!d!", ulReserved1->getData());
		if (rgPropVals)
		{
			for (const auto& prop : rgPropVals->Props())
			{
				addChild(prop);
			}
		}
	}

	void RecipientRowStream::parse()
	{
		m_cVersion = blockT<DWORD>::parse(parser);
		m_cRowCount = blockT<DWORD>::parse(parser);

		if (*m_cRowCount)
		{
			if (*m_cRowCount < _MaxEntriesSmall)
			{
				m_lpAdrEntry.reserve(*m_cRowCount);
				for (DWORD i = 0; i < *m_cRowCount; i++)
				{
					m_lpAdrEntry.emplace_back(block::parse<ADRENTRYStruct>(parser, false));
				}
			}
		}
	}

	void RecipientRowStream::parseBlocks()
	{
		setText(L"Recipient Row Stream");
		addChild(m_cVersion, L"cVersion = %1!d!", m_cVersion->getData());
		addChild(m_cRowCount, L"cRowCount = %1!d!", m_cRowCount->getData());
		if (!m_lpAdrEntry.empty())
		{
			auto i = DWORD{};
			for (const auto& entry : m_lpAdrEntry)
			{
				addChild(entry, L"Row %1!d!", i);
				i++;
			}
		}
	}
} // namespace smartview