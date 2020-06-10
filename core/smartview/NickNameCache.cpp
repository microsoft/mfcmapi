#include <core/stdafx.h>
#include <core/smartview/NickNameCache.h>

namespace smartview
{
	SRowStruct::SRowStruct(const std::shared_ptr<binaryParser>& parser)
	{
		cValues = blockT<DWORD>::parse(parser);

		if (*cValues)
		{
			if (*cValues < _MaxEntriesSmall)
			{
				lpProps = std::make_shared<PropertiesStruct>();
				lpProps->EnableNickNameParsing();
				lpProps->SetMaxEntries(*cValues);
				lpProps->block::parse(parser, false);
			}
		}
	} // namespace smartview

	void NickNameCache::parse()
	{
		m_Metadata1 = blockBytes::parse(parser, 4);
		m_ulMajorVersion = blockT<DWORD>::parse(parser);
		m_ulMinorVersion = blockT<DWORD>::parse(parser);
		m_cRowCount = blockT<DWORD>::parse(parser);

		if (*m_cRowCount)
		{
			if (*m_cRowCount < _MaxEntriesEnormous)
			{
				m_lpRows.reserve(*m_cRowCount);
				for (DWORD i = 0; i < *m_cRowCount; i++)
				{
					m_lpRows.emplace_back(std::make_shared<SRowStruct>(parser));
				}
			}
		}

		m_cbEI = blockT<DWORD>::parse(parser);
		m_lpbEI = blockBytes::parse(parser, *m_cbEI, _MaxBytes);
		m_Metadata2 = blockBytes::parse(parser, 8);
	}

	void NickNameCache::parseBlocks()
	{
		setText(L"Nickname Cache\r\n");
		addLabeledChild(L"Metadata1 = ", m_Metadata1);

		addChild(m_ulMajorVersion, L"Major Version = %1!d!\r\n", m_ulMajorVersion->getData());
		addChild(m_ulMinorVersion, L"Minor Version = %1!d!\r\n", m_ulMinorVersion->getData());
		addChild(m_cRowCount, L"Row Count = %1!d!", m_cRowCount->getData());

		if (!m_lpRows.empty())
		{
			auto i = DWORD{};
			for (const auto& row : m_lpRows)
			{
				terminateBlock();
				if (i > 0) addBlankLine();
				auto rowBlock = create(L"Row %1!d!\r\n", i);
				addChild(rowBlock);
				rowBlock->addChild(row->cValues, L"cValues = 0x%1!08X! = %1!d!\r\n", row->cValues->getData());
				rowBlock->addChild(row->lpProps);

				i++;
			}
		}

		terminateBlock();
		addBlankLine();

		addLabeledChild(L"Extra Info = ", m_lpbEI);

		addLabeledChild(L"Metadata 2 = ", m_Metadata2);
	}
} // namespace smartview