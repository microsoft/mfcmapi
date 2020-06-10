#pragma once
#include <core/smartview/block/binaryParser.h>
#include <core/utility/strings.h>

namespace smartview
{
	constexpr ULONG _MaxBytes = 0xFFFF;
	constexpr ULONG _MaxDepth = 25;
	constexpr ULONG _MaxEID = 500;
	constexpr ULONG _MaxEntriesSmall = 500;
	constexpr ULONG _MaxEntriesLarge = 1000;
	constexpr ULONG _MaxEntriesExtraLarge = 1500;
	constexpr ULONG _MaxEntriesEnormous = 10000;

	class block
	{
	public:
		block() = default;
		block(const block&) = delete;
		block& operator=(const block&) = delete;

		static std::shared_ptr<block> create();
		template <typename... Args>
		static std::shared_ptr<block> create(size_t size, size_t offset, const std::wstring& _text, Args... args)
		{
			auto ret = create();
			ret->setSize(size);
			ret->setOffset(offset);
			ret->setText(strings::formatmessage(_text.c_str(), args...));
			return ret;
		}

		template <typename... Args> static std::shared_ptr<block> create(const std::wstring& _text, Args... args)
		{
			auto ret = create();
			ret->setText(strings::formatmessage(_text.c_str(), args...));
			return ret;
		}

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

		std::wstring toString();

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

		void addHeader(const std::wstring& _text);
		template <typename... Args> void addHeader(const std::wstring& _text, Args... args)
		{
			addHeader(strings::formatmessage(_text.c_str(), args...));
		}

		template <typename... Args> void setText(const std::wstring& _text, Args... args)
		{
			text = strings::formatmessage(_text.c_str(), args...);
		}

		// Add a block as a child
		void addChild(const std::shared_ptr<block>& child)
		{
			if (child && child->isSet())
			{
				children.push_back(child);
			}
		}

		void addChild(const std::shared_ptr<block>& child, const std::wstring& _text)
		{
			if (child && child->isSet())
			{
				child->text = _text;
				children.push_back(child);
			}
		}

		template <typename... Args>
		void addChild(const std::shared_ptr<block>& child, const std::wstring& _text, Args... args)
		{
			if (child && child->isSet())
			{
				child->text = strings::formatmessage(_text.c_str(), args...);
				children.push_back(child);
			}
		}

		void addLabeledChild(const std::wstring& _text, const std::shared_ptr<block>& _block);

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

		void addBlankLine();

		bool hasData() const noexcept { return !text.empty() || !children.empty(); }

		virtual void parse(const std::shared_ptr<binaryParser>& binaryParser, bool _enableJunk)
		{
			parse(binaryParser, 0, _enableJunk);
		}

		virtual void parse(const std::shared_ptr<binaryParser>& binaryParser, size_t cbBin, bool _enableJunk)
		{
			m_Parser = binaryParser;
			m_Parser->setCap(cbBin);
			enableJunk = _enableJunk;
			ensureParsed();
			m_Parser->clearCap();
		}

		template <typename T>
		static std::shared_ptr<T> parse(const std::shared_ptr<binaryParser>& binaryParser, bool _enableJunk)
		{
			return parse<T>(binaryParser, 0, _enableJunk);
		}

		template <typename T>
		static std::shared_ptr<T>
		parse(const std::shared_ptr<binaryParser>& binaryParser, size_t cbBin, bool _enableJunk)
		{
			auto ret = std::make_shared<T>();
			ret->block::parse(binaryParser, cbBin, _enableJunk);
			return ret;
		}

	protected:
		void ensureParsed();
		std::shared_ptr<binaryParser> m_Parser;
		bool parsed{false};
		bool enableJunk{true};

	private:
		std::wstring toStringInternal() const;
		// Consume binaryParser and populate members(which may also inherit from block)
		virtual void parse() = 0;
		// (optional) Stitches block submembers into a tree vis children member
		virtual void parseBlocks(){};

		size_t offset{};
		size_t cb{};
		ULONG source{};
		std::wstring text;
		std::vector<std::shared_ptr<block>> children;
		bool blank{false};
	};
} // namespace smartview