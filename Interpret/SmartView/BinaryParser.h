#pragma once
#include <Interpret/SmartView/block/block.h>
#include <Interpret/SmartView/block/blockT.h>
#include <Interpret/SmartView/block/blockBytes.h>
#include <Interpret/SmartView/block/blockStringA.h>
#include <Interpret/SmartView/block/blockStringW.h>

namespace smartview
{
	// CBinaryParser - helper class for parsing binary data without
	// worrying about whether you've run off the end of your buffer.
	class CBinaryParser
	{
	public:
		CBinaryParser() {}
		CBinaryParser(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin) { Init(cbBin, lpBin); }
		void Init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin)
		{
			m_Bin = lpBin && cbBin ? std::vector<BYTE>(lpBin, lpBin + cbBin) : std::vector<BYTE>{};
			m_Offset = 0;
		}

		bool Empty() const { return m_Bin.empty(); }
		void Advance(size_t cbAdvance) { m_Offset += cbAdvance; }
		void Rewind() { m_Offset = 0; }
		size_t GetCurrentOffset() const { return m_Offset; }
		const BYTE* GetCurrentAddress() const { return m_Bin.data() + m_Offset; }
		// Moves the parser to an offset obtained from GetCurrentOffset
		void SetCurrentOffset(size_t stOffset) { m_Offset = stOffset; }

		// If we're before the end of the buffer, return the count of remaining bytes
		// If we're at or past the end of the buffer, return 0
		// If we're before the beginning of the buffer, return 0
		size_t CBinaryParser::RemainingBytes() const { return m_Offset > m_Bin.size() ? 0 : m_Bin.size() - m_Offset; }

		template <typename T> blockT<T> Get()
		{
			auto ret = blockT<T>();
			// TODO: Consider what a failure block really looks like
			if (!CheckRemainingBytes(sizeof T)) return ret;

			ret.setOffset(m_Offset);
			// TODO: Can we remove this cast?
			ret.setData(*reinterpret_cast<const T*>(GetCurrentAddress()));
			ret.setSize(sizeof T);
			m_Offset += sizeof T;
			return ret;
		}

		blockStringA GetStringA(size_t cchChar = -1)
		{
			auto ret = blockStringA();
			if (cchChar == -1)
			{
				cchChar =
					strnlen_s(reinterpret_cast<LPCSTR>(GetCurrentAddress()), (m_Bin.size() - m_Offset) / sizeof CHAR) +
					1;
			}

			if (!cchChar || !CheckRemainingBytes(sizeof CHAR * cchChar)) return ret;

			ret.setOffset(m_Offset);
			ret.setData(
				strings::RemoveInvalidCharactersA(std::string(reinterpret_cast<LPCSTR>(GetCurrentAddress()), cchChar)));
			ret.setSize(sizeof CHAR * cchChar);
			m_Offset += sizeof CHAR * cchChar;
			return ret;
		}

		blockStringW GetStringW(size_t cchChar = -1)
		{
			auto ret = blockStringW();
			if (cchChar == -1)
			{
				cchChar =
					wcsnlen_s(
						reinterpret_cast<LPCWSTR>(GetCurrentAddress()), (m_Bin.size() - m_Offset) / sizeof WCHAR) +
					1;
			}

			if (!cchChar || !CheckRemainingBytes(sizeof WCHAR * cchChar)) return ret;

			ret.setOffset(m_Offset);
			ret.setData(strings::RemoveInvalidCharactersW(
				std::wstring(reinterpret_cast<LPCWSTR>(GetCurrentAddress()), cchChar)));
			ret.setSize(sizeof WCHAR * cchChar);
			m_Offset += sizeof WCHAR * cchChar;
			return ret;
		}

		blockBytes GetBYTES(size_t cbBytes, size_t cbMaxBytes = -1)
		{
			// TODO: Should we track when the returned byte length is less than requested?
			auto ret = blockBytes();
			ret.setOffset(m_Offset);

			if (cbBytes && CheckRemainingBytes(cbBytes) && (cbMaxBytes == -1 || cbBytes <= cbMaxBytes))
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
		bool CheckRemainingBytes(size_t cbBytes) const { return cbBytes <= RemainingBytes(); }
		std::vector<BYTE> m_Bin;
		size_t m_Offset{};
	};
} // namespace smartview