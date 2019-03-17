#include <core/stdafx.h>
#include <core/smartview/TaskAssigners.h>

namespace smartview
{
	TaskAssigner::TaskAssigner(std::shared_ptr<binaryParser>& parser)
	{
		cbAssigner.parse<DWORD>(parser);
		const auto ulSize = min(cbAssigner, (ULONG) parser->RemainingBytes());
		parser->setCap(ulSize);
		cbEntryID.parse<DWORD>(parser);
		lpEntryID = blockBytes::parse(parser, cbEntryID, _MaxEID);
		szDisplayName = blockStringA::parse(parser);
		wzDisplayName = blockStringW::parse(parser);
		JunkData = blockBytes::parse(parser);
		parser->clearCap();
	}

	void TaskAssigners::Parse()
	{
		m_cAssigners.parse<DWORD>(m_Parser);

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
		addChild(m_cAssigners, L"cAssigners = 0x%1!08X! = %1!d!", m_cAssigners.getData());

		auto i = 0;
		for (const auto& ta : m_lpTaskAssigners)
		{
			terminateBlock();
			addHeader(L"Task Assigner[%1!d!]\r\n", i);
			addChild(ta->cbEntryID, L"\tcbEntryID = 0x%1!08X! = %1!d!\r\n", ta->cbEntryID.getData());
			addHeader(L"\tlpEntryID = ");

			if (!ta->lpEntryID->empty())
			{
				addChild(ta->lpEntryID);
			}

			terminateBlock();
			addChild(ta->szDisplayName, L"\tszDisplayName (ANSI) = %1!hs!\r\n", ta->szDisplayName->c_str());
			addChild(ta->wzDisplayName, L"\tszDisplayName (Unicode) = %1!ws!", ta->wzDisplayName->c_str());

			if (!ta->JunkData->empty())
			{
				terminateBlock();
				addChild(ta->JunkData, L"\tUnparsed Data Size = 0x%1!08X!\r\n", ta->JunkData->size());
				addChild(ta->JunkData);
			}

			i++;
		}
	}
} // namespace smartview