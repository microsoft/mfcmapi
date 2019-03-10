#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockT.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockStringW.h>

namespace smartview
{
	// binaryParser - helper class for parsing binary data without
	// worrying about whether you've run off the end of your buffer.
	class binaryParser
	{
	public:
		binaryParser() = default;
		binaryParser(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin)
		{
			m_Bin = lpBin && cbBin ? std::vector<BYTE>(lpBin, lpBin + cbBin) : std::vector<BYTE>{};
			m_Size = m_Bin.size();
		}
		binaryParser(const binaryParser&) = delete;
		binaryParser& operator=(const binaryParser&) = delete;

		bool empty() const { return m_Offset == m_Size; }
		void advance(size_t cbAdvance) { m_Offset += cbAdvance; }
		void rewind() { m_Offset = 0; }
		size_t GetCurrentOffset() const { return m_Offset; }
		const BYTE* GetCurrentAddress() const { return m_Bin.data() + m_Offset; }
		void SetCurrentOffset(size_t stOffset) { m_Offset = stOffset; }
		void setCap(size_t cap)
		{
			m_Sizes.push(m_Size);
			if (cap != 0 && m_Offset + cap < m_Size)
			{
				m_Size = m_Offset + cap;
			}
		}
		void clearCap()
		{
			if (m_Sizes.empty())
			{
				m_Size = m_Bin.size();
			}
			else
			{
				m_Size = m_Sizes.top();
				m_Sizes.pop();
			}
		}

		// If we're before the end of the buffer, return the count of remaining bytes
		// If we're at or past the end of the buffer, return 0
		// If we're before the beginning of the buffer, return 0
		size_t RemainingBytes() const { return m_Offset > m_Size ? 0 : m_Size - m_Offset; }
		bool CheckRemainingBytes(size_t cbBytes) const { return cbBytes <= RemainingBytes(); }

		template <typename T> blockT<T> Get()
		{
			// TODO: Consider what a failure block really looks like
			if (!CheckRemainingBytes(sizeof T)) return {};

			auto ret = blockT<T>();
			ret.setOffset(m_Offset);
			// TODO: Can we remove this cast?
			ret.setData(*reinterpret_cast<const T*>(GetCurrentAddress()));
			ret.setSize(sizeof T);
			m_Offset += sizeof T;
			return ret;
		}

		blockStringW GetStringW(size_t cchChar = -1)
		{
			if (cchChar == static_cast<size_t>(-1))
			{
				cchChar =
					wcsnlen_s(reinterpret_cast<LPCWSTR>(GetCurrentAddress()), (m_Size - m_Offset) / sizeof WCHAR) + 1;
			}

			auto ret = blockStringW();
			if (cchChar && CheckRemainingBytes(sizeof WCHAR * cchChar))
			{
				ret.setOffset(m_Offset);
				ret.setData(strings::RemoveInvalidCharactersW(
					std::wstring(reinterpret_cast<LPCWSTR>(GetCurrentAddress()), cchChar)));
				ret.setSize(sizeof WCHAR * cchChar);
				m_Offset += sizeof WCHAR * cchChar;
			}

			return ret;
		}

		blockBytes GetBYTES(size_t cbBytes, size_t cbMaxBytes = -1)
		{
			// TODO: Should we track when the returned byte length is less than requested?
			auto ret = blockBytes();
			ret.setOffset(m_Offset);

			if (cbBytes && CheckRemainingBytes(cbBytes) &&
				(cbMaxBytes == static_cast<size_t>(-1) || cbBytes <= cbMaxBytes))
			{
				ret.setData(std::vector<BYTE>{const_cast<LPBYTE>(GetCurrentAddress()),
											  const_cast<LPBYTE>(GetCurrentAddress() + cbBytes)});
				m_Offset += cbBytes;
			}

			// Important that we set our size after getting data, because we may not have gotten the requested byte length
			ret.setSize(ret.size() * sizeof(BYTE));
			return ret;
		}

		blockBytes GetRemainingData() { return GetBYTES(RemainingBytes()); }

	private:
		std::vector<BYTE> m_Bin;
		size_t m_Offset{};
		size_t m_Size{}; // When uncapped, this is m_Bin.size(). When capped, this is our artificial capped size.
		std::stack<size_t> m_Sizes;
	};
} // namespace smartview