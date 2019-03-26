#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/BinaryParser.h>

namespace smartview
{
	template <typename T, typename V = T> class blockT : public block
	{
	public:
		blockT() = default;
		blockT(const blockT&) = delete;
		blockT& operator=(const blockT&) = delete;
		blockT(const std::shared_ptr<binaryParser>& parser) { parse<T>(parser); }

		// Mimic type T
		void setData(const T& _data) { data = _data; }
		T getData() const { return data; }
		operator T&() { return data; }
		operator T() const { return data; }

		// Parse as type U, but store into type T
		template <typename U> void parse(const std::shared_ptr<binaryParser>& parser)
		{
			// TODO: Consider what a failure block really looks like
			if (!parser->CheckRemainingBytes(sizeof U)) return;

			setOffset(parser->GetCurrentOffset());
			// TODO: Can we remove this cast?
			setData(*reinterpret_cast<const U*>(parser->GetCurrentAddress()));
			setSize(sizeof U);
			parser->advance(sizeof U);
		}

		// Parse as type V, but store into type T
		static std::shared_ptr<blockT<T>> parse(const std::shared_ptr<binaryParser>& parser)
		{
			auto ret = std::make_shared<blockT<T>>();
			ret->parse<V>(parser);
			return ret;
		}

	private:
		T data{};
	};

	template <typename T> std::shared_ptr<blockT<T>> emptyT() { return std::make_shared<blockT<T>>(); }
} // namespace smartview