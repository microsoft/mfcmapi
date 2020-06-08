#include <core/stdafx.h>
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/scratchBlock.h>

namespace smartview
{
	void block::ensureParsed()
	{
		if (parsed || m_Parser->empty()) return;
		parsed = true; // parse can unset this if needed
		const auto startOffset = m_Parser->getOffset();

		parse();
		parseBlocks();

		if (this->hasData() && enableJunk && m_Parser->getSize())
		{
			auto junkData = blockBytes::parse(m_Parser, m_Parser->getSize());
			terminateBlock();
			addChild(header(L"Unparsed data size = 0x%1!08X!\r\n", junkData->size()));
			addChild(junkData);
		}

		const auto endOffset = m_Parser->getOffset();

		// Ensure we are tagged with offset and size so all top level blocks get proper highlighting
		setOffset(startOffset);
		setSize(endOffset - startOffset);
	}
} // namespace smartview