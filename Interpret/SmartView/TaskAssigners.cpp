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
				CBinaryParser AssignerParser(ulSize, m_Parser.GetCurrentAddress());
				taskAssigner.cbEntryID = AssignerParser.Get<DWORD>();
				taskAssigner.lpEntryID = AssignerParser.GetBYTES(taskAssigner.cbEntryID, _MaxEID);
				taskAssigner.szDisplayName = AssignerParser.GetStringA();
				taskAssigner.wzDisplayName = AssignerParser.GetStringW();
				taskAssigner.JunkData = AssignerParser.GetRemainingData();

				m_Parser.advance(ulSize);
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
			addBlankLine();
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

			addBlankLine();
			addBlock(
				m_lpTaskAssigners[i].szDisplayName,
				L"\tszDisplayName (ANSI) = %1!hs!",
				m_lpTaskAssigners[i].szDisplayName.c_str());
			addBlankLine();
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