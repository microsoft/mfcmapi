#include <core/stdafx.h>
#include <core/smartview/RecipientRowStream.h>
#include <core/smartview/PropertiesStruct.h>

namespace smartview
{
	ADRENTRYStruct::ADRENTRYStruct(std::shared_ptr<binaryParser> parser)
	{
		cValues = parser->Get<DWORD>();
		ulReserved1 = parser->Get<DWORD>();

		if (cValues && cValues < _MaxEntriesSmall)
		{
			rgPropVals.SetMaxEntries(cValues);
			rgPropVals.parse(parser, false);
		}
	}

	void RecipientRowStream::Parse()
	{
		m_cVersion = m_Parser->Get<DWORD>();
		m_cRowCount = m_Parser->Get<DWORD>();

		if (m_cRowCount && m_cRowCount < _MaxEntriesSmall)
		{
			m_lpAdrEntry.reserve(m_cRowCount);
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				m_lpAdrEntry.emplace_back(m_Parser);
			}
		}
	}

	void RecipientRowStream::ParseBlocks()
	{
		setRoot(L"Recipient Row Stream\r\n");
		addBlock(m_cVersion, L"cVersion = %1!d!\r\n", m_cVersion.getData());
		addBlock(m_cRowCount, L"cRowCount = %1!d!\r\n", m_cRowCount.getData());
		if (!m_lpAdrEntry.empty() && m_cRowCount)
		{
			addBlankLine();
			auto i = DWORD{};
			for (const auto& entry : m_lpAdrEntry)
			{
				terminateBlock();
				addHeader(L"Row %1!d!\r\n", i++);
				addBlock(entry.cValues, L"cValues = 0x%1!08X! = %1!d!\r\n", entry.cValues.getData());
				addBlock(entry.ulReserved1, L"ulReserved1 = 0x%1!08X! = %1!d!\r\n", entry.ulReserved1.getData());
				addBlock(entry.rgPropVals.getBlock());
			}
		}
	}
} // namespace smartview