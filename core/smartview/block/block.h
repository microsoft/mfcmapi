#pragma once
#include <core/utility/strings.h>

namespace smartview
{
	class block
	{
	public:
		block() = default;
		explicit block(std::wstring _text) : text(std::move(_text)) {}
		block& operator=(const block&) = delete;

		virtual bool set() const { return true; }
		const std::wstring& getText() const { return text; }
		const std::vector<std::shared_ptr<block>>& getChildren() const { return children; }
		bool isHeader() const { return cb == 0 && offset == 0; }

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

		void addChild(block& child, const std::wstring& _text)
		{
			if (!child.set()) return;
			child.text = _text;
			children.push_back(std::make_shared<block>(child));
		}

		void addChild(std::shared_ptr<block> child, const std::wstring& _text)
		{
			if (!set() || !child) return;
			child->text = _text;
			children.push_back(child);
		}

		template <typename... Args> void addChild(block& child, const std::wstring& _text, Args... args)
		{
			if (child.set()) addChild(child, strings::formatmessage(_text.c_str(), args...));
		}

		template <typename... Args> void addChild(std::shared_ptr<block> child, const std::wstring& _text, Args... args)
		{
			if (child->set()) addChild(child, strings::formatmessage(_text.c_str(), args...));
		}

		// Add a block as a child
		void addChild(block& child)
		{
			if (child.set()) addChild(child, child.toStringInternal());
		}

		void addChild(std::shared_ptr<block> child)
		{
			if (child->set()) addChild(child, child->toStringInternal());
		}

		// Copy a block into this block with text
		template <typename... Args> void setBlock(block& _data, const std::wstring& _text, Args... args)
		{
			if (!_data.set()) return;
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
		virtual std::wstring toStringInternal() const { return text; }
		std::wstring text;
		std::vector<std::shared_ptr<block>> children;
		bool blank{false};
	};

	template <class T> std::shared_ptr<T> empty() { return std::make_shared<T>(); }
} // namespace smartview