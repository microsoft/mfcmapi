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

		void init(size_t _cb, _In_count_(_cb) const BYTE* _bin)
		{
			parser = std::make_shared<binaryParser>(_cb, _bin);
			parsed = false;
			enableJunk = true;
		}

		// Getters and setters
		// Get the text for just this block
		const std::wstring& getText() const noexcept { return text; }
		// Get the text for this and all child blocks
		std::wstring toString();
		template <typename... Args> void setText(const std::wstring& _text, Args... args)
		{
			text = strings::formatmessage(_text.c_str(), args...);
		}
		void setText(const std::wstring& _text) { text = _text.c_str(); }

		const std::vector<std::shared_ptr<block>>& getChildren() const noexcept { return children; }
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

		bool isSet() const noexcept { return parsed; }
		bool isHeader() const noexcept { return cb == 0 && offset == 0; }
		bool hasData() const noexcept { return !text.empty() || !children.empty(); }

		// Add child blocks of various types
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

		void addHeader(const std::wstring& _text);
		template <typename... Args> void addHeader(const std::wstring& _text, Args... args)
		{
			addHeader(strings::formatmessage(_text.c_str(), args...));
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

		// Static create functions returns a non parsing block
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

		// Static parse functions return a parsing block based on a binaryParser
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

		// Non-static parse functions actually do the parsing
		void parse(const std::shared_ptr<binaryParser>& binaryParser, bool _enableJunk)
		{
			parse(binaryParser, 0, _enableJunk);
		}

		void parse(const std::shared_ptr<binaryParser>& binaryParser, size_t cbBin, bool _enableJunk)
		{
			parser = binaryParser;
			parser->setCap(cbBin);
			enableJunk = _enableJunk;
			ensureParsed();
			parser->clearCap();
		}

	protected:
		void ensureParsed();
		std::shared_ptr<binaryParser> parser;
		bool parsed{false};
		bool enableJunk{true};

	private:
		std::vector<std::wstring> toStringsInternal() const;
		// Consume binaryParser and populate members(which may also inherit from block)
		virtual void parse() = 0;
		// (optional) Stitches block submembers into a tree vis children member
		virtual void parseBlocks(){};

		size_t offset{};
		size_t cb{};
		ULONG source{};
		std::wstring text;
		std::vector<std::shared_ptr<block>> children;
	};
} // namespace smartview