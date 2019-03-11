#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/BinaryParser.h>

namespace smartview
{
	template <class T> class blockT : public block
	{
	public:
		blockT() = default;
		blockT(const blockT&) = delete;
		blockT& operator=(const blockT&) = delete;
		blockT(std::shared_ptr<binaryParser>& parser) { parse<T>(parser); }

		void setData(const T& _data) { data = _data; }
		T getData() const { return data; }
		operator T&() { return data; }
		operator T() const { return data; }

		template <class U> operator blockT<U>() const
		{
			auto _data = blockT<U>(block(*this));
			_data.setData(data);
			return _data;
		}

		// Parse as type U, but store into type T
		template <class U> void parse(std::shared_ptr<binaryParser>& parser)
		{
			// TODO: Consider what a failure block really looks like
			if (!parser->CheckRemainingBytes(sizeof U)) return;

			setOffset(parser->GetCurrentOffset());
			// TODO: Can we remove this cast?
			setData(*reinterpret_cast<const U*>(parser->GetCurrentAddress()));
			setSize(sizeof U);
			parser->advance(sizeof U);
		}

	private:
		T data{};
	};
} // namespace smartview