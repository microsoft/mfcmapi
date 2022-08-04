#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/binaryParser.h>

namespace smartview
{
	class blockBytes : public block
	{
	public:
		blockBytes() = default;
		blockBytes(const blockBytes&) = delete;
		blockBytes& operator=(const blockBytes&) = delete;

		// Mimic std::vector<BYTE>
		operator const std::vector<BYTE>&() const noexcept { return _data; }
		size_t size() const noexcept { return _data.size(); }
		bool empty() const noexcept { return _data.empty(); }
		const BYTE* data() const noexcept { return _data.data(); }

		static std::shared_ptr<blockBytes>
		parse(const std::shared_ptr<binaryParser>& parser, size_t _cbBytes, size_t _cbMaxBytes = -1)
		{
			auto ret = std::make_shared<blockBytes>();
			ret->parser = parser;
			ret->enableJunk = false;
			ret->cbBytes = _cbBytes;
			ret->cbMaxBytes = _cbMaxBytes;
			ret->ensureParsed();
			return ret;
		}

		std::wstring toTextString(bool bMultiLine) const
		{
			return strings::StripCharacter(strings::BinToTextStringW(_data, bMultiLine), L'\0');
		}

		std::wstring toHexString(bool bMultiLine) const { return strings::BinToHexString(_data, bMultiLine); }

		bool equal(size_t _cb, const BYTE* _bin) const noexcept
		{
			if (_cb != _data.size()) return false;
			return memcmp(_data.data(), _bin, _cb) == 0;
		}

	private:
		void parse() override
		{
			parsed = false;
			if (cbBytes && parser->checkSize(cbBytes) &&
				(cbMaxBytes == static_cast<size_t>(-1) || cbBytes <= cbMaxBytes))
			{
				_data = std::vector<BYTE>{
					const_cast<LPBYTE>(parser->getAddress()), const_cast<LPBYTE>(parser->getAddress() + cbBytes)};
				parser->advance(cbBytes);
				setText(toHexString(true));
				parsed = true;
			}
		};

		// TODO: Would it be better to hold the parser and size/offset data and build this as needed?
		std::vector<BYTE> _data;

		size_t cbBytes{};
		size_t cbMaxBytes{};
	};

	inline std::shared_ptr<blockBytes> emptyBB() { return std::make_shared<blockBytes>(); }
} // namespace smartview