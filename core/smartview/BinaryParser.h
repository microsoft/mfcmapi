#pragma once

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

	private:
		std::vector<BYTE> m_Bin;
		size_t m_Offset{};
		size_t m_Size{}; // When uncapped, this is m_Bin.size(). When capped, this is our artificial capped size.
		std::stack<size_t> m_Sizes;
	};
} // namespace smartview