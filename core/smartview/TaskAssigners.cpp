#include <core/stdafx.h>
#include <core/smartview/TaskAssigners.h>

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
		m_cAssigners = blockT<DWORD>::parse(parser);

		if (*m_cAssigners && *m_cAssigners < _MaxEntriesSmall)
		{
			m_lpTaskAssigners.reserve(*m_cAssigners);
			for (DWORD i = 0; i < *m_cAssigners; i++)
			{
				m_lpTaskAssigners.emplace_back(std::make_shared<TaskAssigner>(parser));
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
			auto taBlock = create(L"Task Assigner[%1!d!]", i);
			addChild(taBlock);
			taBlock->addChild(ta->cbEntryID, L"cbEntryID = 0x%1!08X! = %1!d!", ta->cbEntryID->getData());
			taBlock->addLabeledChild(L"lpEntryID =", ta->lpEntryID);
			taBlock->addChild(ta->szDisplayName, L"szDisplayName (ANSI) = %1!hs!", ta->szDisplayName->c_str());
			taBlock->addChild(ta->wzDisplayName, L"szDisplayName (Unicode) = %1!ws!", ta->wzDisplayName->c_str());

			if (!ta->JunkData->empty())
			{
				taBlock->addChild(ta->JunkData, L"Unparsed Data Size = 0x%1!08X!", ta->JunkData->size());
				taBlock->addChild(ta->JunkData);
			}

			i++;
		}
	}
} // namespace smartview