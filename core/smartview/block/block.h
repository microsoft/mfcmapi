#pragma once
#include <core/utility/strings.h>

namespace smartview
{
	class block
	{
	public:
		block() = default;

		const std::wstring& getText() const { return text; }
		const std::vector<std::shared_ptr<block>>& getChildren() const { return children; }
		explicit block(std::wstring _text) : text(std::move(_text)) {}
		bool isHeader() const { return cb == 0 && offset == 0; }
		void clear()
		{
			offset = 0;
			cb = 0;
			source = 0;

			text.clear();
			children.clear();
			blank = false;
		}

		virtual std::wstring ToString() const
		{
			std::vector<std::wstring> items;
			items.reserve(children.size() + 1);
			items.push_back(blank ? L"\r\n" : text);

			for (const auto& item : children)
			{
				items.emplace_back(item->ToString());
			}

			return strings::join(items, strings::emptystring);
		}

		size_t getSize() const { return cb; }
		void setSize(size_t _size) { cb = _size; }
		size_t getOffset() const { return offset; }
		void setOffset(size_t _offset) { offset = _offset; }
		ULONG getSource() const { return source; }
		void setSource(ULONG _source)
		{
			source = _source;
			for (auto& child : children)
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

		void addBlock(block& child, const std::wstring& _text)
		{
			child.text = _text;
			children.push_back(std::make_shared<block>(child));
		}

		template <typename... Args> void addBlock(block& child, const std::wstring& _text, Args... args)
		{
			child.text = strings::formatmessage(_text.c_str(), args...);
			children.push_back(std::make_shared<block>(child));
		}

		// Add a block as a child
		void addBlock(block& child)
		{
			child.text = child.ToStringInternal();
			children.emplace_back(std::make_shared<block>(child));
		}

		// Copy a block into this block with text
		template <typename... Args> void setBlock(block& _data, const std::wstring& _text, Args... args)
		{
			text = strings::formatmessage(_text.c_str(), args...);
			children = _data.children;
		}

		void terminateBlock()
		{
			if (children.size() >= 1)
			{
				children.back()->terminateBlock();
			}
			else
			{
				text = strings::ensureCRLF(text);
			}
		}

		void addBlankLine()
		{
			auto child = std::make_shared<block>();
			child->blank = true;
			children.push_back(child);
		}

		bool hasData() const { return !text.empty() || !children.empty(); }

	protected:
		size_t offset{};
		size_t cb{};
		ULONG source{};

	private:
		virtual std::wstring ToStringInternal() const { return text; }
		std::wstring text;
		std::vector<std::shared_ptr<block>> children;
		bool blank{false};
	};
} // namespace smartview