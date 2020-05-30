#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/block/binaryParser.h>

namespace smartview
{
	template <typename T, typename SOURCE_T = T> class blockT : public smartViewParser
	{
	public:
		blockT() = default;
		blockT(const blockT&) = delete;
		blockT& operator=(const blockT&) = delete;
		blockT(const T& _data, size_t _size, size_t _offset) noexcept
		{
			set = true;
			data = _data;
			setSize(_size);
			setOffset(_offset);
		}

		bool isSet() const noexcept override { return set; }

		// Mimic type T
		void setData(const T& _data) noexcept { data = _data; }
		T getData() const noexcept { return data; }
		operator T&() noexcept { return data; }
		operator T() const noexcept { return data; }

		// Fill out object of type T, reading from type SOURCE_U
		template <typename SOURCE_U> void parse(const std::shared_ptr<binaryParser>& parser) noexcept
		{
			static constexpr size_t sizeU = sizeof SOURCE_U;
			// TODO: Consider what a failure block really looks like
			if (!parser->checkSize(sizeU)) return;

			setOffset(parser->getOffset());
			// TODO: Can we remove this cast?
			setData(*reinterpret_cast<const SOURCE_U*>(parser->getAddress()));
			setSize(sizeU);
			parser->advance(sizeU);

			set = true;
		}

		// Build and return object of type T, reading from type SOURCE_T
		// Usage: std::shared_ptr<blockT<T>> tmp = blockT<T, SOURCE_T>::parser(parser);
		static std::shared_ptr<blockT<T>> parse(const std::shared_ptr<binaryParser>& parser)
		{
			auto ret = std::make_shared<blockT<T>>();
			ret->parse<SOURCE_T>(parser);
			return ret;
		}

		static std::shared_ptr<blockT<T>> create(const T& _data, size_t _size, size_t _offset)
		{
			return std::make_shared<blockT<T>>(_data, _size, _offset);
		}

		void parse() override
		{
		};
		void parseBlocks() override{};

	private:
		// Construct directly from a parser
		blockT(const std::shared_ptr<binaryParser>& parser) { parse<T>(parser); }
		T data{};
		bool set{false};
	};

	template <typename T> std::shared_ptr<blockT<T>> emptyT() { return std::make_shared<blockT<T>>(); }
} // namespace smartview