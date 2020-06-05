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

	static std::shared_ptr<scratchBlock> create() { return std::make_shared<scratchBlock>(); }
} // namespace smartview