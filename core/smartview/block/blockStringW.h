#pragma once
#include <core/smartview/block/block.h>

namespace smartview
{
	class blockStringW : public block
	{
	public:
		blockStringW() = default;
		blockStringW(const blockStringW&) = delete;
		blockStringW& operator=(const blockStringW&) = delete;

		// Mimic std::wstring
		operator const std::wstring&() const noexcept { return data; }
		_NODISCARD _Ret_z_ const wchar_t* c_str() const noexcept { return data.c_str(); }
		_NODISCARD std::wstring::size_type length() const noexcept { return data.length(); }
		_NODISCARD bool empty() const noexcept { return data.empty(); }

		static std::shared_ptr<blockStringW> parse(const std::wstring& _data, size_t _size, size_t _offset)
		{
			auto ret = std::make_shared<blockStringW>();
			ret->parsed = true;
			ret->enableJunk = false;
			ret->data = _data;
			ret->setText(_data);
			ret->setSize(_size);
			ret->setOffset(_offset);
			return ret;
		}

		static std::shared_ptr<blockStringW> parse(const std::shared_ptr<binaryParser>& parser, size_t cchChar = -1)
		{
			auto ret = std::make_shared<blockStringW>();
			ret->parser = parser;
			ret->enableJunk = false;
			ret->cchChar = cchChar;
			ret->ensureParsed();
			return ret;
		}

	private:
		void parse() override
		{
			parsed = false;
			if (cchChar == static_cast<size_t>(-1))
			{
				cchChar =
					wcsnlen_s(reinterpret_cast<LPCWSTR>(parser->getAddress()), (parser->getSize()) / sizeof WCHAR) + 1;
			}

			if (cchChar && parser->checkSize(sizeof WCHAR * cchChar))
			{
				data = strings::RemoveInvalidCharactersW(
					std::wstring(reinterpret_cast<LPCWSTR>(parser->getAddress()), cchChar));
				parser->advance(sizeof WCHAR * cchChar);
				setText(data);
				parsed = true;
			}
		};

		std::wstring data;

		size_t cchChar{};
	};

	inline std::shared_ptr<blockStringW> emptySW() { return std::make_shared<blockStringW>(); }
} // namespace smartview