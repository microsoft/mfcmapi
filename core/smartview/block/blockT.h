#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/block/binaryParser.h>

namespace smartview
{
	template <typename T> class blockT : public smartViewParser
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

		void setSet(bool _set) { set = _set; }
		// Build and return object of type T, reading from type SOURCE_T
		// Usage: std::shared_ptr<blockT<T>> tmp = blockT<T, SOURCE_T>::parser(parser);
		static std::shared_ptr<blockT<T>> parse(const std::shared_ptr<binaryParser>& parser)
		{
			auto ret = std::make_shared<blockT<T>>();
			ret->m_Parser = parser;
			ret->ensureParsed();
			return ret;
		}

		template <typename U> static std::shared_ptr<blockT<T>> parse(const std::shared_ptr<binaryParser>& parser)
		{
			auto ret = std::make_shared<blockT<T>>();
			static constexpr size_t sizeU = sizeof U;
			if (!parser->checkSize(sizeU)) return ret;

			ret->setOffset(parser->getOffset());
			ret->setData(*reinterpret_cast<const U*>(parser->getAddress()));
			ret->setSize(sizeU);
			parser->advance(sizeU);

			ret->setSet(true);
			return ret;
		}

		static std::shared_ptr<blockT<T>> create(const T& _data, size_t _size, size_t _offset)
		{
			return std::make_shared<blockT<T>>(_data, _size, _offset);
		}

		void parse() override
		{
			if (!m_Parser->checkSize(sizeof T)) return;

			data = *reinterpret_cast<const T*>(m_Parser->getAddress());
			m_Parser->advance(sizeof T);
			set = true;
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