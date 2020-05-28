#include <core/stdafx.h>
#include <core/smartview/RuleAction.h>
#include <core/smartview/RestrictionStruct.h>
#include <core/interpret/guid.h>

namespace smartview
{
	void RuleAction::Init(bool bExtended) noexcept { m_bExtended = bExtended; }

	void RuleAction::parse()
	{
	}

	void RuleAction::parseBlocks()
	{
		setRoot(m_bExtended ? L"Extended Rule Action\r\n" : L"Rule Action\r\n");
	}
} // namespace smartview