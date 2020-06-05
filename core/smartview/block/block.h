#pragma once
#include <core/smartview/block/binaryParser.h>
#include <core/utility/strings.h>

namespace smartview
{
	class block
	{
	public:
		block() = default;
		explicit block(std::wstring _text) noexcept : text(std::move(_text)) {}
		block(const block&) = delete;
		block& operator=(const block&) = delete;

		virtual bool isSet() const noexcept { return true; }
		const std::wstring& getText() const noexcept { return text; }
		const std::vector<std::shared_ptr<block>>& getChildren() const noexcept { return children; }
		bool isHeader() const noexcept { return cb == 0 && offset == 0; }

		virtual std::wstring toString() const
		{
			std::vector<std::wstring> items;
			items.reserve(children.size() + 1);
			items.push_back(blank ? L"\r\n" : text);

			for (const auto& item : children)
			{
				items.emplace_back(item->toString());
			}

			return strings::join(items, strings::emptystring);
		}

		size_t getSize() const noexcept { return cb; }
		void setSize(size_t _size) noexcept { cb = _size; }
		size_t getOffset() const noexcept { return offset; }
		void setOffset(size_t _offset) noexcept { offset = _offset; }
		ULONG getSource() const noexcept { return source; }
		void setSource(ULONG _source)
		{
			source = _source;
			for (const auto& child : children)
			{
				child->setSource(_source);
			}
		}

		template <typename... Args> void addHeader(const std::wstring& _text, Args... args)
		{
			children.emplace_back(std::make_shared<block>(strings::formatmessage(_text.c_str(), args...)));
		}

		template <typename... Args> void setText(const std::wstring& _text, Args... args)
		{
			text = strings::formatmessage(_text.c_str(), args...);
		}

		// Add a block as a child
		void addChild(const std::shared_ptr<block>& child)
		{
			if (child && child->isSet()) addChild(child, child->toStringInternal());
		}

		void addChild(const std::shared_ptr<block>& child, const std::wstring& _text)
		{
			if (!child || !child->isSet()) return;
			child->text = _text;
			children.push_back(child);
		}

		template <typename... Args>
		void addChild(const std::shared_ptr<block>& child, const std::wstring& _text, Args... args)
		{
			if (child->isSet()) addChild(child, strings::formatmessage(_text.c_str(), args...));
		}

		void terminateBlock()
		{
			if (children.empty())
			{
				text = strings::ensureCRLF(text);
			}
			else
			{
				children.back()->terminateBlock();
			}
		}

		void setBlank(bool _blank) { blank = _blank; }

		bool hasData() const noexcept { return !text.empty() || !children.empty(); }

	protected:
		std::shared_ptr<binaryParser> m_Parser;
		bool parsed{false};
		bool enableJunk{true};

	private:
		virtual std::wstring toStringInternal() const { return text; }

		size_t offset{};
		size_t cb{};
		ULONG source{};
		std::wstring text;
		std::vector<std::shared_ptr<block>> children;
		bool blank{false};
	};
} // namespace smartview