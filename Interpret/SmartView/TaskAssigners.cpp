#include <StdAfx.h>
#include <Interpret/SmartView/TaskAssigners.h>

namespace smartview
{
	TaskAssigners::TaskAssigners() {}

	void TaskAssigners::Parse()
	{
		m_cAssigners = m_Parser.GetBlock<DWORD>();

		if (m_cAssigners.getData() && m_cAssigners.getData() < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_cAssigners.getData(); i++)
			{
				TaskAssigner taskAssigner;
				taskAssigner.cbAssigner = m_Parser.GetBlock<DWORD>();
				const auto ulSize = min(taskAssigner.cbAssigner.getData(), (ULONG) m_Parser.RemainingBytes());
				CBinaryParser AssignerParser(ulSize, m_Parser.GetCurrentAddress());
				taskAssigner.cbEntryID = AssignerParser.GetBlock<DWORD>();
				taskAssigner.lpEntryID = AssignerParser.GetBlockBYTES(taskAssigner.cbEntryID.getData(), _MaxEID);
				taskAssigner.szDisplayName = AssignerParser.GetBlockStringA();
				taskAssigner.wzDisplayName = AssignerParser.GetBlockStringW();
				taskAssigner.JunkData = AssignerParser.GetBlockRemainingData();

				m_Parser.Advance(ulSize);
				m_lpTaskAssigners.push_back(taskAssigner);
			}
		}
	}

	_Check_return_ std::wstring TaskAssigners::ToStringInternal()
	{
		std::wstring szTaskAssigners;

		szTaskAssigners += strings::formatmessage(IDS_TASKASSIGNERSHEADER, m_cAssigners);

		for (DWORD i = 0; i < m_lpTaskAssigners.size(); i++)
		{
			szTaskAssigners += strings::formatmessage(IDS_TASKASSIGNEREID, i, m_lpTaskAssigners[i].cbEntryID);

			if (!m_lpTaskAssigners[i].lpEntryID.getData().empty())
			{
				szTaskAssigners += strings::BinToHexString(m_lpTaskAssigners[i].lpEntryID.getData(), true);
			}

			szTaskAssigners += strings::formatmessage(
				IDS_TASKASSIGNERNAME,
				m_lpTaskAssigners[i].szDisplayName.getData().c_str(),
				m_lpTaskAssigners[i].wzDisplayName.getData().c_str());

			if (!m_lpTaskAssigners[i].JunkData.getData().empty())
			{
				szTaskAssigners +=
					strings::formatmessage(IDS_TASKASSIGNERJUNKDATA, m_lpTaskAssigners[i].JunkData.getData().size());
				szTaskAssigners += strings::BinToHexString(m_lpTaskAssigners[i].JunkData.getData(), true);
			}
		}

		return szTaskAssigners;
	}
}