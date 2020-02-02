#pragma once
#include <core/smartview/block/block.h>

namespace smartview
{
	class blockStringA : public block
	{
	public:
		blockStringA() = default;
		blockStringA(const blockStringA&) = delete;
		blockStringA& operator=(const blockStringA&) = delete;

		virtual bool isSet() const override { return set; }

		// Mimic std::string
		operator const std::string&() const { return data; }
		_NODISCARD _Ret_z_ const char* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::string::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }
		std::wstring toWstring() const noexcept { return strings::stringTowstring(data); }

		blockStringA(const std::shared_ptr<binaryParser>& parser, size_t cchChar = -1)
		{
			if (cchChar == static_cast<size_t>(-1))
			{
				cchChar =
					strnlen_s(reinterpret_cast<LPCSTR>(parser->getAddress()), (parser->getSize()) / sizeof CHAR) + 1;
			}

			if (cchChar && parser->checkSize(sizeof CHAR * cchChar))
			{
				setOffset(parser->getOffset());
				data = strings::RemoveInvalidCharactersA(
					std::string(reinterpret_cast<LPCSTR>(parser->getAddress()), cchChar));
				setSize(sizeof CHAR * cchChar);
				parser->advance(sizeof CHAR * cchChar);
				set = true;
			}
		}

		blockStringA(const std::string& _data, size_t _size, size_t _offset)
		{
			set = true;
			data = _data;
			setSize(_size);
			setOffset(_offset);
		}

		static std::shared_ptr<blockStringA> parse(const std::shared_ptr<binaryParser>& parser, size_t cchChar = -1)
		{
			return std::make_shared<blockStringA>(parser, cchChar);
		}

	private:
		std::wstring toStringInternal() const override { return strings::stringTowstring(data); }
		std::string data;
		bool set{false};
	};

	inline std::shared_ptr<blockStringA> emptySA() { return std::make_shared<blockStringA>(); }
} // namespace smartview