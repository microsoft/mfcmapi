#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/BinaryParser.h>

namespace smartview
{
	template <typename T, typename SOURCE_T = T> class blockT : public block
	{
	public:
		blockT() = default;
		blockT(const blockT&) = delete;
		blockT& operator=(const blockT&) = delete;
		// Construct directly from a parser
		blockT(const std::shared_ptr<binaryParser>& parser) { parse<T>(parser); }

		// Mimic type T
		void setData(const T& _data) { data = _data; }
		T getData() const { return data; }
		operator T&() { return data; }
		operator T() const { return data; }

		// Fill out object of type T, reading from type SOURCE_U
		template <typename SOURCE_U> void parse(const std::shared_ptr<binaryParser>& parser)
		{
			static const size_t sizeU = sizeof SOURCE_U;
			// TODO: Consider what a failure block really looks like
			if (!parser->CheckRemainingBytes(sizeU)) return;

			setOffset(parser->GetCurrentOffset());
			// TODO: Can we remove this cast?
			setData(*reinterpret_cast<const SOURCE_U*>(parser->GetCurrentAddress()));
			setSize(sizeU);
			parser->advance(sizeU);
		}

		// Build and return object of type T, reading from type SOURCE_T
		// Usage: std::shared_ptr<blockT<T>> tmp = blockT<T, SOURCE_T>::parser(parser);
		static std::shared_ptr<blockT<T>> parse(const std::shared_ptr<binaryParser>& parser)
		{
			auto ret = std::make_shared<blockT<T>>();
			ret->parse<SOURCE_T>(parser);
			return ret;
		}

	private:
		T data{};
	};

	template <typename T> std::shared_ptr<blockT<T>> emptyT() { return std::make_shared<blockT<T>>(); }
} // namespace smartview