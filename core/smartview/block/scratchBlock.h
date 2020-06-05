#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/utility/strings.h>

namespace smartview
{
	class scratchBlock : public smartViewParser
	{
	public:
		scratchBlock() { parsed = true; }

		template <typename... Args>
		static std::shared_ptr<scratchBlock> create(size_t size, size_t offset, const std::wstring& _text, Args... args)
		{
			auto ret = smartview::create();
			ret->setSize(size);
			ret->setOffset(offset);
			ret->setText(strings::formatmessage(_text.c_str(), args...));
			return ret;
		}

		template <typename... Args> static std::shared_ptr<scratchBlock> create(const std::wstring& _text, Args... args)
		{
			auto ret = smartview::create();
			ret->setText(strings::formatmessage(_text.c_str(), args...));
			return ret;
		}

	private:
		void parse() override{};
	};

	inline std::shared_ptr<scratchBlock> create() { return std::make_shared<scratchBlock>(); }
	inline std::shared_ptr<scratchBlock> blankLine()
	{
		auto ret = create();
		ret->setBlank(true);
		return ret;
	}
	inline std::shared_ptr<scratchBlock> labeledBlock(const std::wstring& _text, const std::shared_ptr<block>& _block)
	{
		if (_block->isSet())
		{
			auto node = create();
			node->setText(_text);
			node->setOffset(_block->getOffset());
			node->setSize(_block->getSize());
			node->addChild(_block);
			node->terminateBlock();
			return node;
		}

		return {};
	}

} // namespace smartview