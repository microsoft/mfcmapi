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

		virtual bool isSet() const { return set; }

		// Mimic std::string
		operator const std::string&() const { return data; }
		_NODISCARD _Ret_z_ const char* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::string::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }
		std::wstring toWstring() const noexcept { return strings::stringTowstring(data); }

		blockStringA(std::shared_ptr<binaryParser> parser, size_t cchChar = -1)
		{
			set = true;
			if (cchChar == static_cast<size_t>(-1))
			{
				cchChar = strnlen_s(
							  reinterpret_cast<LPCSTR>(parser->GetCurrentAddress()),
							  (parser->RemainingBytes()) / sizeof CHAR) +
						  1;
			}

			if (cchChar && parser->CheckRemainingBytes(sizeof CHAR * cchChar))
			{
				setOffset(parser->GetCurrentOffset());
				data = strings::RemoveInvalidCharactersA(
					std::string(reinterpret_cast<LPCSTR>(parser->GetCurrentAddress()), cchChar));
				setSize(sizeof CHAR * cchChar);
				parser->advance(sizeof CHAR * cchChar);
			}
		}

		blockStringA(const std::string& _data, size_t _size, size_t _offset)
		{
			set = true;
			data = _data;
			setSize(_size);
			setOffset(_offset);
		}

		static std::shared_ptr<blockStringA> parse(std::shared_ptr<binaryParser> parser, size_t cchChar = -1)
		{
			return std::make_shared<blockStringA>(parser, cchChar);
		}

	private:
		std::string data;
		bool set{false};
	};

	inline std::shared_ptr<blockStringA> emptySA() { return std::make_shared<blockStringA>(); }
} // namespace smartview