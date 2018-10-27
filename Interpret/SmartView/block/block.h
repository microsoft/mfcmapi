#pragma once

namespace smartview
{
	class block
	{
	public:
		block() : offset(0), cb(0), text(L""), header(true) {}

		std::wstring getText() const { return text; }
		const std::vector<block>& getChildren() const { return children; }

		virtual std::wstring ToString() const
		{
			std::vector<std::wstring> items;
			items.reserve(children.size() + 1);
			items.push_back(text);
			for (const auto& item : children)
			{
				items.push_back(item.ToString());
			}

			return strings::join(items, strings::emptystring);
		}

		size_t getSize() const { return cb; }
		void setSize(size_t _size) { cb = _size; }
		size_t getOffset() const { return offset; }
		void setOffset(size_t _offset) { offset = _offset; }
		template <typename... Args> void addHeader(const std::wstring& _text, Args... args)
		{
			children.emplace_back(block(strings::formatmessage(_text.c_str(), args...)));
		}

		void addBlock(const block& child, const std::wstring& _text)
		{
			auto block = child;
			block.text = _text;
			children.push_back(block);
		}

		template <typename... Args> void addBlock(const block& child, const std::wstring& _text, Args... args)
		{
			auto block = child;
			block.text = strings::formatmessage(_text.c_str(), args...);
			children.push_back(block);
		}

		void addBlock(const block& child)
		{
			auto block = child;
			block.text = child.ToStringInternal();
			children.push_back(block);
		}

		void addLine() { addHeader(L"\r\n"); }
		bool hasData() const { return !text.empty() || !children.empty(); }

	protected:
		size_t offset;
		size_t cb;

	private:
		explicit block(std::wstring _text) : offset(0), cb(0), text(std::move(_text)), header(true) {}
		virtual std::wstring ToStringInternal() const { return text; }
		std::wstring text;
		std::vector<block> children;
		bool header;
	};
} // namespace smartview