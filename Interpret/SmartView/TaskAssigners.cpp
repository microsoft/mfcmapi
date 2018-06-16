#include <StdAfx.h>
#include <Interpret/SmartView/TaskAssigners.h>

namespace smartview
{
	TaskAssigners::TaskAssigners() { m_cAssigners = 0; }

	void TaskAssigners::Parse()
	{
		m_cAssigners = m_Parser.Get<DWORD>();

		if (m_cAssigners && m_cAssigners < _MaxEntriesSmall)
		{
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

			if (!m_lpTaskAssigners[i].lpEntryID.empty())
			{
				szTaskAssigners += strings::BinToHexString(m_lpTaskAssigners[i].lpEntryID, true);
			}

			szTaskAssigners += strings::formatmessage(
				IDS_TASKASSIGNERNAME,
				m_lpTaskAssigners[i].szDisplayName.c_str(),
				m_lpTaskAssigners[i].wzDisplayName.c_str());

			if (!m_lpTaskAssigners[i].JunkData.empty())
			{
				szTaskAssigners +=
					strings::formatmessage(IDS_TASKASSIGNERJUNKDATA, m_lpTaskAssigners[i].JunkData.size());
				szTaskAssigners += strings::BinToHexString(m_lpTaskAssigners[i].JunkData, true);
			}
		}

		return szTaskAssigners;
	}
}