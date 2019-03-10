#include <core/stdafx.h>
#include <core/smartview/TaskAssigners.h>

namespace smartview
{
	TaskAssigner::TaskAssigner(std::shared_ptr<binaryParser>& parser)
	{
		cbAssigner = parser->Get<DWORD>();
		const auto ulSize = min(cbAssigner, (ULONG) parser->RemainingBytes());
		parser->setCap(ulSize);
		cbEntryID = parser->Get<DWORD>();
		lpEntryID = parser->GetBYTES(cbEntryID, _MaxEID);
		szDisplayName.parse(parser);
		wzDisplayName.parse(parser);
		JunkData = parser->GetRemainingData();
		parser->clearCap();
	}

	void TaskAssigners::Parse()
	{
		m_cAssigners = m_Parser->Get<DWORD>();

		if (m_cAssigners && m_cAssigners < _MaxEntriesSmall)
		{
			m_lpTaskAssigners.reserve(m_cAssigners);
			for (DWORD i = 0; i < m_cAssigners; i++)
			{
				m_lpTaskAssigners.emplace_back(std::make_shared<TaskAssigner>(m_Parser));
			}
		}
	}

	void TaskAssigners::ParseBlocks()
	{
		setRoot(L"Task Assigners: \r\n");
		addBlock(m_cAssigners, L"cAssigners = 0x%1!08X! = %1!d!", m_cAssigners.getData());

		for (DWORD i = 0; i < m_lpTaskAssigners.size(); i++)
		{
			terminateBlock();
			addHeader(L"Task Assigner[%1!d!]\r\n", i);
			addBlock(
				m_lpTaskAssigners[i]->cbEntryID,
				L"\tcbEntryID = 0x%1!08X! = %1!d!\r\n",
				m_lpTaskAssigners[i]->cbEntryID.getData());
			addHeader(L"\tlpEntryID = ");

			if (!m_lpTaskAssigners[i]->lpEntryID.empty())
			{
				addBlock(m_lpTaskAssigners[i]->lpEntryID);
			}

			terminateBlock();
			addBlock(
				m_lpTaskAssigners[i]->szDisplayName,
				L"\tszDisplayName (ANSI) = %1!hs!\r\n",
				m_lpTaskAssigners[i]->szDisplayName.c_str());
			addBlock(
				m_lpTaskAssigners[i]->wzDisplayName,
				L"\tszDisplayName (Unicode) = %1!ws!",
				m_lpTaskAssigners[i]->wzDisplayName.c_str());

			if (!m_lpTaskAssigners[i]->JunkData.empty())
			{
				terminateBlock();
				addBlock(
					m_lpTaskAssigners[i]->JunkData,
					L"\tUnparsed Data Size = 0x%1!08X!\r\n",
					m_lpTaskAssigners[i]->JunkData.size());
				addBlock(m_lpTaskAssigners[i]->JunkData);
			}
		}
	}
} // namespace smartview