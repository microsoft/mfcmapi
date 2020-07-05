#include <core/stdafx.h>
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/scratchBlock.h>

namespace smartview
{
	std::shared_ptr<block> block::create() { return std::make_shared<scratchBlock>(); }

	void block::addHeader(const std::wstring& _text) { addChild(create(_text)); }

	void block::addLabeledChild(const std::wstring& _text, const std::shared_ptr<block>& _block)
	{
		if (_block->isSet())
		{
			auto node = create();
			node->setText(_text);
			node->setOffset(_block->getOffset());
			node->setSize(_block->getSize());
			node->addChild(_block);
			addChild(node);
		}
	}

	void block::addSubHeader(const std::wstring& _text)
	{
		auto node = create();
		node->setText(_text);
		node->setOffset(getOffset());
		node->setSize(getSize());
		addChild(node);
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
			addLabeledChild(strings::formatmessage(L"Unparsed data size = 0x%1!08X!", junkData->size()), junkData);
		}

		const auto endOffset = parser->getOffset();

		// Ensure we are tagged with offset and size so all top level blocks get proper highlighting
		setOffset(startOffset);
		setSize(endOffset - startOffset);
	}

	std::vector<std::wstring> tabStrings(const std::vector<std::wstring>& elems, bool usePipes)
	{
		if (elems.empty()) return {};

		std::vector<std::wstring> strings;
		strings.reserve(elems.size());
		auto iter = elems.begin();
		for (const auto& elem : elems)
		{
			if (usePipes)
			{
				strings.emplace_back(L"|\t" + elem);
			}
			else
			{
				strings.emplace_back(L"\t" + elem);
			}
		}

		return strings;
	}

	std::vector<std::wstring> block::toStringsInternal() const
	{
		std::vector<std::wstring> strings;
		strings.reserve(children.size() + 1);
		if (!text.empty()) strings.push_back(text + L"\r\n");

		for (const auto& child : children)
		{
			auto childStrings = child->toStringsInternal();
			if (!text.empty()) childStrings = tabStrings(childStrings, usePipes());
			strings.insert(std::end(strings), std::begin(childStrings), std::end(childStrings));
		}

		return strings;
	}

	std::wstring block::toString()
	{
		ensureParsed();

		auto strings = toStringsInternal();
		auto parsedString = strings::trimWhitespace(strings::join(strings, strings::emptystring));

		// If we built a string with embedded nulls in it, replace them with dots.
		std::replace_if(
			parsedString.begin(), parsedString.end(), [](const WCHAR& chr) { return chr == L'\0'; }, L'.');

		return parsedString;
	}
} // namespace smartview