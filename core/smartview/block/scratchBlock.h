#pragma once
#include <core/smartview/block/block.h>
#include <core/utility/strings.h>

namespace smartview
{
	class scratchBlock : public block
	{
	public:
		scratchBlock() noexcept { parsed = true; }

	private:
		void parse() override{};
	};

	inline std::shared_ptr<scratchBlock> create() { return std::make_shared<scratchBlock>(); }

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
} // namespace smartview