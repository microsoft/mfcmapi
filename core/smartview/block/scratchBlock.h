#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/utility/strings.h>

namespace smartview
{
	class scratchBlock : public smartViewParser
	{
	public:
		static std::shared_ptr<scratchBlock> create()
		{
			auto ret = std::make_shared<scratchBlock>();
			ret->parsed = true;
			return ret;
		}

		template <typename... Args>
		static std::shared_ptr<scratchBlock> create(size_t size, size_t offset, const std::wstring& _text, Args... args)
		{
			auto ret = create();
			ret->setSize(size);
			ret->setOffset(offset);
			ret->setText(strings::formatmessage(_text.c_str(), args...));
			return ret;
		}

		template <typename... Args> static std::shared_ptr<scratchBlock> create(const std::wstring& _text, Args... args)
		{
			auto ret = create();
			ret->setText(strings::formatmessage(_text.c_str(), args...));
			return ret;
		}

	private:
		void parse() override{};
	};
} // namespace smartview