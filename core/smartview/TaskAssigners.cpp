#include <core/stdafx.h>
#include <core/smartview/TaskAssigners.h>
#include <core/smartview/block/scratchBlock.h>

namespace smartview
{
	TaskAssigner::TaskAssigner(const std::shared_ptr<binaryParser>& parser)
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

	void TaskAssigners::parse()
	{
		m_cAssigners = blockT<DWORD>::parse(m_Parser);

		if (*m_cAssigners && *m_cAssigners < _MaxEntriesSmall)
		{
			m_lpTaskAssigners.reserve(*m_cAssigners);
			for (DWORD i = 0; i < *m_cAssigners; i++)
			{
				m_lpTaskAssigners.emplace_back(std::make_shared<TaskAssigner>(m_Parser));
			}
		}
	}

	void TaskAssigners::parseBlocks()
	{
		setText(L"Task Assigners: \r\n");
		addChild(m_cAssigners, L"cAssigners = 0x%1!08X! = %1!d!", m_cAssigners->getData());

		auto i = 0;
		for (const auto& ta : m_lpTaskAssigners)
		{
			terminateBlock();
			addHeader(L"Task Assigner[%1!d!]\r\n", i);
			addChild(ta->cbEntryID, L"\tcbEntryID = 0x%1!08X! = %1!d!\r\n", ta->cbEntryID->getData());
			addLabeledChild(L"\tlpEntryID = ", ta->lpEntryID);
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