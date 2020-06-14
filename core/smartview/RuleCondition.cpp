#include <core/stdafx.h>
#include <core/smartview/RuleCondition.h>
#include <core/smartview/RestrictionStruct.h>

namespace smartview
{
	void PropertyName::parse()
	{
		Kind = blockT<BYTE>::parse(parser);
		Guid = blockT<GUID>::parse(parser);
		if (*Kind == MNID_ID)
		{
			LID = blockT<DWORD>::parse(parser);
		}
		else if (*Kind == MNID_STRING)
		{
			NameSize = blockT<BYTE>::parse(parser);
			Name = blockStringW::parse(parser, *NameSize / sizeof(WCHAR));
		}
	}

	void RuleCondition::Init(bool bExtended) noexcept { m_bExtended = bExtended; }

	void RuleCondition::parse()
	{
		m_NamedPropertyInformation = block::parse<NamedPropertyInformation>(parser, 0, false);
		m_lpRes = std::make_shared<RestrictionStruct>(true, m_bExtended);
		m_lpRes->block::parse(parser, false);
	}

	void RuleCondition::parseBlocks()
	{
		setText(m_bExtended ? L"Extended Rule Condition\r\n" : L"Rule Condition\r\n");
		addChild(m_NamedPropertyInformation);
		if (m_lpRes && m_lpRes->hasData())
		{
			terminateBlock();
			addChild(m_lpRes);
		}
	}
} // namespace smartview