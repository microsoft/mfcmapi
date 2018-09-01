#include <StdAfx.h>
#include <Interpret/SmartView/RecipientRowStream.h>
#include <Interpret/String.h>
#include <Interpret/SmartView/PropertyStruct.h>

namespace smartview
{
	RecipientRowStream::RecipientRowStream() {}

	void RecipientRowStream::Parse()
	{
		m_cVersion = m_Parser.GetBlock<DWORD>();
		m_cRowCount = m_Parser.GetBlock<DWORD>();

		if (m_cRowCount && m_cRowCount < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				auto entry = ADRENTRYStruct{};
				entry.cValues = m_Parser.GetBlock<DWORD>();
				entry.ulReserved1 = m_Parser.GetBlock<DWORD>();

				if (entry.cValues && entry.cValues < _MaxEntriesSmall)
				{
					entry.rgPropVals.Init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
					entry.rgPropVals.DisableJunkParsing();
					entry.rgPropVals.SetMaxEntries(entry.cValues);
					entry.rgPropVals.EnsureParsed();
					m_Parser.Advance(entry.rgPropVals.GetCurrentOffset());
				}

				m_lpAdrEntry.push_back(entry);
			}
		}
	}

	_Check_return_ std::wstring RecipientRowStream::ToStringInternal()
	{
		auto szRecipientRowStream =
			strings::formatmessage(IDS_RECIPIENTROWSTREAMHEADER, m_cVersion.getData(), m_cRowCount.getData());
		if (m_lpAdrEntry.size() && m_cRowCount)
		{
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				szRecipientRowStream += strings::formatmessage(
					IDS_RECIPIENTROWSTREAMROW,
					i,
					m_lpAdrEntry[i].cValues.getData(),
					m_lpAdrEntry[i].ulReserved1.getData());

				// TODO: Use blocks
				szRecipientRowStream += m_lpAdrEntry[i].rgPropVals.ToString();
			}
		}

		return szRecipientRowStream;
	}
}