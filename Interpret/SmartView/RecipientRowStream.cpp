#include <StdAfx.h>
#include <Interpret/SmartView/RecipientRowStream.h>
#include <Interpret/String.h>
#include <Interpret/SmartView/PropertiesStruct.h>

namespace smartview
{
	void RecipientRowStream::Parse()
	{
		m_cVersion = m_Parser.Get<DWORD>();
		m_cRowCount = m_Parser.Get<DWORD>();

		if (m_cRowCount && m_cRowCount < _MaxEntriesSmall)
		{
			m_lpAdrEntry.reserve(m_cRowCount);
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				auto entry = ADRENTRYStruct{};
				entry.cValues = m_Parser.Get<DWORD>();
				entry.ulReserved1 = m_Parser.Get<DWORD>();

				if (entry.cValues && entry.cValues < _MaxEntriesSmall)
				{
					entry.rgPropVals.init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
					entry.rgPropVals.DisableJunkParsing();
					entry.rgPropVals.SetMaxEntries(entry.cValues);
					entry.rgPropVals.EnsureParsed();
					m_Parser.advance(entry.rgPropVals.GetCurrentOffset());
				}

				m_lpAdrEntry.push_back(entry);
			}
		}
	}

	void RecipientRowStream::ParseBlocks()
	{
		addHeader(L"Recipient Row Stream\r\n");
		addBlock(m_cVersion, L"cVersion = %1!d!\r\n", m_cVersion.getData());
		addBlock(m_cRowCount, L"cRowCount = %1!d!\r\n", m_cRowCount.getData());
		if (m_lpAdrEntry.size() && m_cRowCount)
		{
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				addHeader(L"\r\nRow %1!d!\r\n", i);
				addBlock(
					m_lpAdrEntry[i].cValues, L"cValues = 0x%1!08X! = %1!d!\r\n", m_lpAdrEntry[i].cValues.getData());
				addBlock(
					m_lpAdrEntry[i].ulReserved1,
					L"ulReserved1 = 0x%1!08X! = %1!d!\r\n",
					m_lpAdrEntry[i].ulReserved1.getData());

				addBlock(m_lpAdrEntry[i].rgPropVals.getBlock());
			}
		}
	}
} // namespace smartview