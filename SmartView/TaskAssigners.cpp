#include "stdafx.h"
#include "TaskAssigners.h"
#include "String.h"

TaskAssigners::TaskAssigners()
{
	m_cAssigners = 0;
	m_lpTaskAssigners = nullptr;
}

TaskAssigners::~TaskAssigners()
{
	if (m_cAssigners && m_lpTaskAssigners)
	{
		for (DWORD i = 0; i < m_cAssigners; i++)
		{
			delete[] m_lpTaskAssigners[i].lpEntryID;
			delete[] m_lpTaskAssigners[i].szDisplayName;
			delete[] m_lpTaskAssigners[i].wzDisplayName;
			delete[] m_lpTaskAssigners[i].JunkData;
		}
	}

	delete[] m_lpTaskAssigners;
}

void TaskAssigners::Parse()
{
	m_Parser.GetDWORD(&m_cAssigners);

	if (m_cAssigners && m_cAssigners < _MaxEntriesSmall)
		m_lpTaskAssigners = new TaskAssigner[m_cAssigners];

	if (m_lpTaskAssigners)
	{
		memset(m_lpTaskAssigners, 0, sizeof(TaskAssigner)* m_cAssigners);
		for (DWORD i = 0; i < m_cAssigners; i++)
		{
			m_Parser.GetDWORD(&m_lpTaskAssigners[i].cbAssigner);
			auto ulSize = min(m_lpTaskAssigners[i].cbAssigner, (ULONG)m_Parser.RemainingBytes());
			CBinaryParser AssignerParser(ulSize, m_Parser.GetCurrentAddress());
			AssignerParser.GetDWORD(&m_lpTaskAssigners[i].cbEntryID);
			AssignerParser.GetBYTES(m_lpTaskAssigners[i].cbEntryID, _MaxEID, &m_lpTaskAssigners[i].lpEntryID);
			AssignerParser.GetStringA(&m_lpTaskAssigners[i].szDisplayName);
			AssignerParser.GetStringW(&m_lpTaskAssigners[i].wzDisplayName);
			m_lpTaskAssigners[i].JunkDataSize = AssignerParser.GetRemainingData(&m_lpTaskAssigners[i].JunkData);

			m_Parser.Advance(ulSize);
		}
	}
}

_Check_return_ wstring TaskAssigners::ToStringInternal()
{
	wstring szTaskAssigners;

	szTaskAssigners += formatmessage(IDS_TASKASSIGNERSHEADER,
		m_cAssigners);

	if (m_cAssigners && m_lpTaskAssigners)
	{
		for (DWORD i = 0; i < m_cAssigners; i++)
		{
			szTaskAssigners += formatmessage(IDS_TASKASSIGNEREID,
				i,
				m_lpTaskAssigners[i].cbEntryID);

			if (m_lpTaskAssigners[i].lpEntryID)
			{
				SBinary sBin = { 0 };
				sBin.cb = m_lpTaskAssigners[i].cbEntryID;
				sBin.lpb = m_lpTaskAssigners[i].lpEntryID;
				szTaskAssigners += BinToHexString(&sBin, true);
			}

			szTaskAssigners += formatmessage(IDS_TASKASSIGNERNAME,
				m_lpTaskAssigners[i].szDisplayName,
				m_lpTaskAssigners[i].wzDisplayName);

			if (m_lpTaskAssigners[i].JunkDataSize)
			{
				szTaskAssigners += formatmessage(IDS_TASKASSIGNERJUNKDATA,
					m_lpTaskAssigners[i].JunkDataSize);
				SBinary sBin = { 0 };
				sBin.cb = static_cast<ULONG>(m_lpTaskAssigners[i].JunkDataSize);
				sBin.lpb = m_lpTaskAssigners[i].JunkData;
				szTaskAssigners += BinToHexString(&sBin, true);
			}
		}
	}

	return szTaskAssigners;
}