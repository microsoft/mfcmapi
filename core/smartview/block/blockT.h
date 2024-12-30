#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/binaryParser.h>

namespace smartview
{
	template <typename T> class blockT : public block
	{
	public:
		blockT() = default;
		blockT(const blockT&) = delete;
		blockT& operator=(const blockT&) = delete;
		blockT(const T& _data, size_t _size, size_t _offset) noexcept
		{
			parsed = true;
			data = _data;
			setSize(_size);
			setOffset(_offset);
		}

		// Mimic type T
		void setData(const T& _data) noexcept { data = _data; }
		T getData() const noexcept { return data; }
		const T* getDataAddress() const noexcept { return &data; }
		operator T&() noexcept { return data; }
		operator T() const noexcept { return data; }

		static std::shared_ptr<blockT<T>> parse(const std::shared_ptr<binaryParser>& parser)
		{
			auto ret = std::make_shared<blockT<T>>();
			ret->parser = parser;
			ret->ensureParsed();
			return ret;
		}

		// Build and return object of type T, reading from type U
		// Usage: std::shared_ptr<blockT<T>> tmp = blockT<T>::parser<U>(parser);
		template <typename U> static std::shared_ptr<blockT<T>> parse(const std::shared_ptr<binaryParser>& parser)
		{
			if (!parser->checkSize(sizeof(U))) return std::make_shared<blockT<T>>();

			const U _data = *reinterpret_cast<const U*>(parser->getAddress());
			const auto offset = parser->getOffset();
			parser->advance(sizeof(U));
			return create(_data, sizeof(U), offset);
		}

		static std::shared_ptr<blockT<T>> create(const T& _data, size_t _size, size_t _offset)
		{
			const auto ret = std::make_shared<blockT<T>>(_data, _size, _offset);
			ret->parsed = true;
			return ret;
		}

	private:
		void parse() override
		{
			parsed = false;
			if (!parser->checkSize(sizeof(T))) return;

			data = *reinterpret_cast<const T*>(parser->getAddress());
			parser->advance(sizeof(T));
			parsed = true;
		};

		// Construct directly from a parser
		blockT(const std::shared_ptr<binaryParser>& parser) { parse<T>(parser); }
		T data{};
	};

	template <typename T> std::shared_ptr<blockT<T>> emptyT() { return std::make_shared<blockT<T>>(); }
} // namespace smartview