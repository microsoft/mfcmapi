#pragma once
#include <core/smartview/block/binaryParser.h>
#include <core/utility/strings.h>

namespace smartview
{
#define _MaxBytes 0xFFFF
#define _MaxDepth 25
#define _MaxEID 500
#define _MaxEntriesSmall 500
#define _MaxEntriesLarge 1000
#define _MaxEntriesExtraLarge 1500
#define _MaxEntriesEnormous 10000

	class block
	{
	public:
		block() = default;
		explicit block(std::wstring _text) noexcept : text(std::move(_text)) {}
		block(const block&) = delete;
		block& operator=(const block&) = delete;

		void init(size_t _cb, _In_count_(_cb) const BYTE* _bin)
		{
			m_Parser = std::make_shared<binaryParser>(_cb, _bin);
			parsed = false;
			enableJunk = true;
		}

		bool isSet() const noexcept { return parsed; }
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
		void ensureParsed();
		std::shared_ptr<binaryParser> m_Parser;
		bool parsed{false};
		bool enableJunk{true};

		// Consume binaryParser and populate (which may also inherit from smarViewParser)
		virtual void parse() = 0;
		// (optional) Stiches smartViewParser submembers into a tree
		virtual void parseBlocks(){};

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