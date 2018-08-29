#include <StdAfx.h>
#include <Interpret/SmartView/NickNameCache.h>
#include <Interpret/String.h>

namespace smartview
{
	NickNameCache::NickNameCache() {}

	void NickNameCache::Parse()
	{
		m_Metadata1 = m_Parser.GetBlockBYTES(4);
		m_ulMajorVersion = m_Parser.GetBlock<DWORD>();
		m_ulMinorVersion = m_Parser.GetBlock<DWORD>();
		m_cRowCount = m_Parser.GetBlock<DWORD>();

		if (m_cRowCount && m_cRowCount < _MaxEntriesEnormous)
		{
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				auto row = SRowStruct{};
				row.cValues = m_Parser.GetBlock<DWORD>();

				if (row.cValues && row.cValues < _MaxEntriesSmall)
				{
					row.lpProps.Init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
					row.lpProps.DisableJunkParsing();
					row.lpProps.EnableNickNameParsing();
					row.lpProps.SetMaxEntries(row.cValues);
					row.lpProps.EnsureParsed();
					m_Parser.Advance(row.lpProps.GetCurrentOffset());
				}

				m_lpRows.push_back(row);
			}
		}

		m_cbEI = m_Parser.GetBlock<DWORD>();
		m_lpbEI = m_Parser.GetBlockBYTES(m_cbEI, _MaxBytes);
		m_Metadata2 = m_Parser.GetBlockBYTES(8);
	}

	_Check_return_ std::wstring NickNameCache::ToStringInternal()
	{
		auto szNickNameCache = strings::formatmessage(IDS_NICKNAMEHEADER);
		szNickNameCache += strings::BinToHexString(m_Metadata1, true);

		szNickNameCache += strings::formatmessage(
			IDS_NICKNAMEROWCOUNT, m_ulMajorVersion.getData(), m_ulMinorVersion.getData(), m_cRowCount.getData());

		if (m_cRowCount && m_lpRows.size())
		{
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				if (i > 0) szNickNameCache += L"\r\n";
				szNickNameCache += strings::formatmessage(IDS_NICKNAMEROWS, i, m_lpRows[i].cValues.getData());

				szNickNameCache += m_lpRows[i].lpProps.ToString();
			}
		}

		szNickNameCache += L"\r\n";
		szNickNameCache += strings::formatmessage(IDS_NICKNAMEEXTRAINFO);
		szNickNameCache += strings::BinToHexString(m_lpbEI, true);

		szNickNameCache += strings::formatmessage(IDS_NICKNAMEFOOTER);
		szNickNameCache += strings::BinToHexString(m_Metadata2, true);

		return szNickNameCache;
	}
}