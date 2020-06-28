#include <core/stdafx.h>
#include <core/smartview/TaskAssigners.h>

namespace smartview
{
	void TaskAssigner::parse()
	{
		cbAssigner = blockT<DWORD>::parse(parser);
		const auto ulSize = min(*cbAssigner, (ULONG) parser->getSize());
		parser->setCap(ulSize);
		cbEntryID = blockT<ULONG>::parse(parser);
		lpEntryID = blockBytes::parse(parser, *cbEntryID, _MaxEID);
		szDisplayName = blockStringA::parse(parser);
		wzDisplayName = blockStringW::parse(parser);
		JunkData = blockBytes::parse(parser, parser->getSize());
		parser->clearCap();
	}

	void TaskAssigner::parseBlocks()
	{
		addChild(cbEntryID, L"cbEntryID = 0x%1!08X! = %1!d!", cbEntryID->getData());
		addLabeledChild(L"lpEntryID", lpEntryID);
		addChild(szDisplayName, L"szDisplayName (ANSI) = %1!hs!", szDisplayName->c_str());
		addChild(wzDisplayName, L"szDisplayName (Unicode) = %1!ws!", wzDisplayName->c_str());

		if (!JunkData->empty())
		{
			addChild(JunkData, L"Unparsed Data Size = 0x%1!08X!", JunkData->size());
			addChild(JunkData);
		}
	}

	void TaskAssigners::parse()
	{
		m_cAssigners = blockT<DWORD>::parse(parser);

		if (*m_cAssigners && *m_cAssigners < _MaxEntriesSmall)
		{
			m_lpTaskAssigners.reserve(*m_cAssigners);
			for (DWORD i = 0; i < *m_cAssigners; i++)
			{
				m_lpTaskAssigners.emplace_back(block::parse<TaskAssigner>(parser, false));
			}
		}
	}

	void TaskAssigners::parseBlocks()
	{
		setText(L"Task Assigners");
		addChild(m_cAssigners, L"cAssigners = 0x%1!08X! = %1!d!", m_cAssigners->getData());

		auto i = 0;
		for (const auto& ta : m_lpTaskAssigners)
		{
			addChild(ta, L"Task Assigner[%1!d!]", i);
			i++;
		}
	}
} // namespace smartview