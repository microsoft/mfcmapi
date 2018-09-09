#include <StdAfx.h>
#include <Interpret/SmartView/block/block.h>
#include <Interpret/SmartView/block/blockBytes.h>
#include <Interpret/String.h>

namespace smartview
{
	void block::addBlockBytes(const blockBytes& child)
	{
		auto _block = child;
		_block.text = strings::BinToHexString(child, true);
		children.push_back(_block);
	}
} // namespace smartview