#include <StdAfx.h>
#include <Interpret/SmartView/TaskAssigners.h>

namespace smartview
{
	void TaskAssigners::Parse()
	{
		m_cAssigners = m_Parser.Get<DWORD>();

		if (m_cAssigners && m_cAssigners < _MaxEntriesSmall)
		{
			m_lpTaskAssigners.reserve(m_cAssigners);
			for (DWORD i = 0; i < m_cAssigners; i++)
			{
				TaskAssigner taskAssigner;
				taskAssigner.cbAssigner = m_Parser.Get<DWORD>();
				const auto ulSize = min(taskAssigner.cbAssigner, (ULONG) m_Parser.RemainingBytes());
				m_Parser.setCap(ulSize);
				taskAssigner.cbEntryID = m_Parser.Get<DWORD>();
				taskAssigner.lpEntryID = m_Parser.GetBYTES(taskAssigner.cbEntryID, _MaxEID);
				taskAssigner.szDisplayName = m_Parser.GetStringA();
				taskAssigner.wzDisplayName = m_Parser.GetStringW();
				taskAssigner.JunkData = m_Parser.GetRemainingData();
				m_Parser.clearCap();

				m_lpTaskAssigners.push_back(taskAssigner);
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
				m_lpTaskAssigners[i].cbEntryID,
				L"\tcbEntryID = 0x%1!08X! = %1!d!\r\n",
				m_lpTaskAssigners[i].cbEntryID.getData());
			addHeader(L"\tlpEntryID = ");

			if (!m_lpTaskAssigners[i].lpEntryID.empty())
			{
				addBlock(m_lpTaskAssigners[i].lpEntryID);
			}

			terminateBlock();
			addBlock(
				m_lpTaskAssigners[i].szDisplayName,
				L"\tszDisplayName (ANSI) = %1!hs!\r\n",
				m_lpTaskAssigners[i].szDisplayName.c_str());
			addBlock(
				m_lpTaskAssigners[i].wzDisplayName,
				L"\tszDisplayName (Unicode) = %1!ws!",
				m_lpTaskAssigners[i].wzDisplayName.c_str());

			if (!m_lpTaskAssigners[i].JunkData.empty())
			{
				terminateBlock();
				addBlock(
					m_lpTaskAssigners[i].JunkData,
					L"\tUnparsed Data Size = 0x%1!08X!\r\n",
					m_lpTaskAssigners[i].JunkData.size());
				addBlock(m_lpTaskAssigners[i].JunkData);
			}
		}
	}
} // namespace smartview