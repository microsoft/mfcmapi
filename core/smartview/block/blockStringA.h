#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/binaryParser.h>

namespace smartview
{
	class blockStringA : public block
	{
	public:
		blockStringA() = default;
		blockStringA(const blockStringA&) = delete;
		blockStringA& operator=(const blockStringA&) = delete;
		void setData(const std::string& _data) { data = _data; }
		operator const std::string&() const { return data; }
		_NODISCARD _Ret_z_ const char* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::string::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

		blockStringA(std::shared_ptr<binaryParser> parser, size_t cchChar = -1) { init(parser, cchChar); }
		void blockStringA::init(std::shared_ptr<binaryParser> parser, size_t cchChar = -1)
		{
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
				setData(strings::RemoveInvalidCharactersA(
					std::string(reinterpret_cast<LPCSTR>(parser->GetCurrentAddress()), cchChar)));
				setSize(sizeof CHAR * cchChar);
				parser->advance(sizeof CHAR * cchChar);
			}
		}

	private:
		std::string data;
	};
} // namespace smartview