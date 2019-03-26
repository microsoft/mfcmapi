#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/BinaryParser.h>

namespace smartview
{
	class blockBytes : public block
	{
	public:
		blockBytes() = default;
		blockBytes(const blockBytes&) = delete;
		blockBytes& operator=(const blockBytes&) = delete;

		void setData(const std::vector<BYTE>& data) { _data = data; }

		// Mimic std::vector<BYTE>
		std::vector<BYTE> getData() const { return _data; }
		operator const std::vector<BYTE>&() const { return _data; }
		size_t size() const noexcept { return _data.size(); }
		bool empty() const noexcept { return _data.empty(); }
		const BYTE* data() const noexcept { return _data.data(); }

		blockBytes(const std::shared_ptr<binaryParser>& parser, size_t cbBytes, size_t cbMaxBytes = -1)
		{
			// TODO: Should we track when the returned byte length is less than requested?
			setOffset(parser->GetCurrentOffset());

			if (cbBytes && parser->CheckRemainingBytes(cbBytes) &&
				(cbMaxBytes == static_cast<size_t>(-1) || cbBytes <= cbMaxBytes))
			{
				setData(std::vector<BYTE>{const_cast<LPBYTE>(parser->GetCurrentAddress()),
										  const_cast<LPBYTE>(parser->GetCurrentAddress() + cbBytes)});
				parser->advance(cbBytes);
			}

			// Important that we set our size after getting data, because we may not have gotten the requested byte length
			setSize(size() * sizeof(BYTE));
		}

		blockBytes(const std::shared_ptr<binaryParser>& parser) : blockBytes(parser, parser->RemainingBytes()) {}

		static std::shared_ptr<blockBytes> parse(const std::shared_ptr<binaryParser>& parser)
		{
			return std::make_shared<blockBytes>(parser, parser->RemainingBytes());
		}

		static std::shared_ptr<blockBytes>
		parse(const std::shared_ptr<binaryParser>& parser, size_t cbBytes, size_t cbMaxBytes = -1)
		{
			return std::make_shared<blockBytes>(parser, cbBytes, cbMaxBytes);
		}

		std::wstring toTextString(bool bMultiLine) const
		{
			return strings::StripCharacter(strings::BinToTextStringW(_data, bMultiLine), L'\0');
		}

	private:
		std::wstring toStringInternal() const override { return strings::BinToHexString(_data, true); }
		std::vector<BYTE> _data;
	};

	inline std::shared_ptr<blockBytes> emptyBB() { return std::make_shared<blockBytes>(); }
} // namespace smartview