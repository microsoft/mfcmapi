#include <StdAfx.h>
#include <Interpret/SmartView/NickNameCache.h>
#include <Interpret/String.h>

namespace smartview
{
	void NickNameCache::Parse()
	{
		m_Metadata1 = m_Parser.GetBYTES(4);
		m_ulMajorVersion = m_Parser.Get<DWORD>();
		m_ulMinorVersion = m_Parser.Get<DWORD>();
		m_cRowCount = m_Parser.Get<DWORD>();

		if (m_cRowCount && m_cRowCount < _MaxEntriesEnormous)
		{
			m_lpRows.reserve(m_cRowCount);
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				auto row = SRowStruct{};
				row.cValues = m_Parser.Get<DWORD>();

				if (row.cValues && row.cValues < _MaxEntriesSmall)
				{
					row.lpProps.EnableNickNameParsing();
					row.lpProps.SetMaxEntries(row.cValues);
					row.lpProps.parse(m_Parser, false);
				}

				m_lpRows.push_back(row);
			}
		}

		m_cbEI = m_Parser.Get<DWORD>();
		m_lpbEI = m_Parser.GetBYTES(m_cbEI, _MaxBytes);
		m_Metadata2 = m_Parser.GetBYTES(8);
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

		if (m_cRowCount && !m_lpRows.empty())
		{
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				terminateBlock();
				if (i > 0) addBlankLine();
				addHeader(L"Row %1!d!\r\n", i);
				addBlock(m_lpRows[i].cValues, L"cValues = 0x%1!08X! = %1!d!\r\n", m_lpRows[i].cValues.getData());

				addBlock(m_lpRows[i].lpProps.getBlock());
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