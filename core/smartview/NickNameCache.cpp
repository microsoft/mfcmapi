#include <core/stdafx.h>
#include <core/smartview/NickNameCache.h>

namespace smartview
{
	SRowStruct::SRowStruct(std::shared_ptr<binaryParser> parser)
	{
		cValues = parser->Get<DWORD>();

		if (cValues)
		{
			if (cValues < _MaxEntriesSmall)
			{
				lpProps.EnableNickNameParsing();
				lpProps.SetMaxEntries(cValues);
				lpProps.parse(parser, false);
			}
		}
	} // namespace smartview

	void NickNameCache::Parse()
	{
		m_Metadata1.parse(m_Parser, 4);
		m_ulMajorVersion = m_Parser->Get<DWORD>();
		m_ulMinorVersion = m_Parser->Get<DWORD>();
		m_cRowCount = m_Parser->Get<DWORD>();

		if (m_cRowCount)
		{
			if (m_cRowCount < _MaxEntriesEnormous)
			{
				m_lpRows.reserve(m_cRowCount);
				for (DWORD i = 0; i < m_cRowCount; i++)
				{
					m_lpRows.emplace_back(std::make_shared<SRowStruct>(m_Parser));
				}
			}
		}

		m_cbEI = m_Parser->Get<DWORD>();
		m_lpbEI.parse(m_Parser, m_cbEI, _MaxBytes);
		m_Metadata2.parse(m_Parser, 8);
	}

	void NickNameCache::ParseBlocks()
	{
		setRoot(L"Nickname Cache\r\n");
		addHeader(L"Metadata1 = ");
		addBlock(m_Metadata1);
		terminateBlock();

		addBlock(m_ulMajorVersion, L"Major Version = %1!d!\r\n", m_ulMajorVersion.getData());
		addBlock(m_ulMinorVersion, L"Minor Version = %1!d!\r\n", m_ulMinorVersion.getData());
		addBlock(m_cRowCount, L"Row Count = %1!d!", m_cRowCount.getData());

		if (!m_lpRows.empty())
		{
			auto i = DWORD{};
			for (const auto& row : m_lpRows)
			{
				terminateBlock();
				if (i > 0) addBlankLine();
				addHeader(L"Row %1!d!\r\n", i++);
				addBlock(row->cValues, L"cValues = 0x%1!08X! = %1!d!\r\n", row->cValues.getData());

				addBlock(row->lpProps.getBlock());
			}
		}

		terminateBlock();
		addBlankLine();
		addHeader(L"Extra Info = ");
		addBlock(m_lpbEI);
		terminateBlock();

		addHeader(L"Metadata 2 = ");
		addBlock(m_Metadata2);
	}
} // namespace smartview