#include <core/stdafx.h>
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/scratchBlock.h>

namespace smartview
{
	std::shared_ptr<block> block::create() { return std::make_shared<scratchBlock>(); }

	void block::addHeader(const std::wstring& _text) { addChild(create(_text)); }

	void block::addBlankLine()
	{
		auto ret = create();
		ret->blank = true;
		children.push_back(ret);
	}

	void block::addLabeledChild(const std::wstring& _text, const std::shared_ptr<block>& _block)
	{
		if (_block->isSet())
		{
			auto node = create();
			node->setText(_text);
			node->setOffset(_block->getOffset());
			node->setSize(_block->getSize());
			node->addChild(_block);
			node->terminateBlock();
			addChild(node);
		}
	}

	void block::ensureParsed()
	{
		if (parsed || !parser || parser->empty()) return;
		parsed = true; // parse can unset this if needed
		const auto startOffset = parser->getOffset();

		parse();
		parseBlocks();

		if (this->hasData() && enableJunk && parser->getSize())
		{
			auto junkData = blockBytes::parse(parser, parser->getSize());
			terminateBlock();
			addHeader(L"Unparsed data size = 0x%1!08X!\r\n", junkData->size());
			addChild(junkData);
		}

		const auto endOffset = parser->getOffset();

		// Ensure we are tagged with offset and size so all top level blocks get proper highlighting
		setOffset(startOffset);
		setSize(endOffset - startOffset);
	}

	std::wstring block::toStringInternal() const
	{
		std::vector<std::wstring> items;
		items.reserve(children.size() + 1);
		items.push_back(blank ? L"\r\n" : text);

		for (const auto& item : children)
		{
			items.emplace_back(item->toStringInternal());
		}

		return strings::join(items, strings::emptystring);
	}

	std::wstring block::toString()
	{
		ensureParsed();

		auto parsedString = strings::trimWhitespace(toStringInternal());

		// If we built a string with embedded nulls in it, replace them with dots.
		std::replace_if(
			parsedString.begin(), parsedString.end(), [](const WCHAR& chr) { return chr == L'\0'; }, L'.');

		return parsedString;
	}
} // namespace smartview