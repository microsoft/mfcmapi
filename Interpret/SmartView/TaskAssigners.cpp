#include <StdAfx.h>
#include <Interpret/SmartView/TaskAssigners.h>

namespace smartview
{
	TaskAssigners::TaskAssigners() {}

	void TaskAssigners::Parse()
	{
		m_cAssigners = m_Parser.GetBlock<DWORD>();

		if (m_cAssigners && m_cAssigners < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_cAssigners; i++)
			{
				TaskAssigner taskAssigner;
				taskAssigner.cbAssigner = m_Parser.GetBlock<DWORD>();
				const auto ulSize = min(taskAssigner.cbAssigner, (ULONG) m_Parser.RemainingBytes());
				CBinaryParser AssignerParser(ulSize, m_Parser.GetCurrentAddress());
				taskAssigner.cbEntryID = AssignerParser.GetBlock<DWORD>();
				taskAssigner.lpEntryID = AssignerParser.GetBlockBYTES(taskAssigner.cbEntryID, _MaxEID);
				taskAssigner.szDisplayName = AssignerParser.GetBlockStringA();
				taskAssigner.wzDisplayName = AssignerParser.GetBlockStringW();
				taskAssigner.JunkData = AssignerParser.GetBlockRemainingData();

				m_Parser.Advance(ulSize);
				m_lpTaskAssigners.push_back(taskAssigner);
			}
		}
	}

	void TaskAssigners::ParseBlocks()
	{
		addHeader(L"Task Assigners: \r\n");
		addBlock(m_cAssigners, L"cAssigners = 0x%1!08X! = %1!d!", m_cAssigners.getData());

		for (DWORD i = 0; i < m_lpTaskAssigners.size(); i++)
		{
			addLine();
			addHeader(L"Task Assigner[%1!d!]\r\n", i);
			addBlock(
				m_lpTaskAssigners[i].cbEntryID,
				L"\tcbEntryID = 0x%1!08X! = %1!d!\r\n",
				m_lpTaskAssigners[i].cbEntryID.getData());
			addHeader(L"\tlpEntryID = ");

			if (!m_lpTaskAssigners[i].lpEntryID.getData().empty())
			{
				addBlockBytes(m_lpTaskAssigners[i].lpEntryID);
			}

			addLine();
			addBlock(
				m_lpTaskAssigners[i].szDisplayName,
				L"\tszDisplayName (ANSI) = %1!hs!",
				m_lpTaskAssigners[i].szDisplayName.getData().c_str());
			addLine();
			addBlock(
				m_lpTaskAssigners[i].wzDisplayName,
				L"\tszDisplayName (Unicode) = %1!ws!",
				m_lpTaskAssigners[i].wzDisplayName.getData().c_str());

			if (!m_lpTaskAssigners[i].JunkData.getData().empty())
			{
				addLine();
				addBlock(
					m_lpTaskAssigners[i].JunkData,
					L"\tUnparsed Data Size = 0x%1!08X!\r\n",
					m_lpTaskAssigners[i].JunkData.getData().size());
				addBlockBytes(m_lpTaskAssigners[i].JunkData);
			}
		}
	}
}