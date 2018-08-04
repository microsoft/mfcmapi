#pragma once

namespace smartview
{
	class block
	{
	public:
		block() : header(true), offset(0), cb(0), text(L"") {}
		block(const std::wstring _text) : header(true), offset(0), cb(0), text(_text) {}
		std::wstring ToString() const
		{
			std::vector<std::wstring> items;
			items.emplace_back(text);
			for (auto item : children)
			{
				items.emplace_back(item.ToString());
			}

			return strings::join(items, strings::emptystring);
		}
		size_t getSize() const { return cb; }
		void setSize(size_t _size) { cb = _size; }
		size_t getOffset() const { return offset; }
		void setOffset(size_t _offset) { offset = _offset; }
		void AddChild(const block& child, const std::wstring _text = L"")
		{
			auto block = child;
			block.text = _text;
			children.push_back(block);
		}

	private:
		size_t offset;
		size_t cb;
		std::wstring text;
		std::vector<block> children;
		bool header;
	};

	class blockBytes : public block
	{
	public:
		blockBytes() {}
		std::vector<BYTE> data;
	};

	template <class T> class blockT : public block
	{
	public:
		blockT() {}
		T data;
	};

	// TODO: Make all of this header only
	// CBinaryParser - helper class for parsing binary data without
	// worrying about whether you've run off the end of your buffer.
	class CBinaryParser
	{
	public:
		CBinaryParser();
		CBinaryParser(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin);

		bool Empty() const;
		void Init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin);
		void Advance(size_t cbAdvance);
		void Rewind();
		size_t GetCurrentOffset() const;
		const BYTE* GetCurrentAddress() const;
		// Moves the parser to an offset obtained from GetCurrentOffset
		void SetCurrentOffset(size_t stOffset);
		size_t RemainingBytes() const;
		template <typename T> T Get()
		{
			if (!CheckRemainingBytes(sizeof T)) return T();
			auto ret = *reinterpret_cast<const T*>(GetCurrentAddress());
			m_Offset += sizeof T;
			return ret;
		}

		template <typename T> blockT<T> GetBlock()
		{
			auto ret = blockT<T>();
			if (!CheckRemainingBytes(sizeof T)) return ret;

			// TODO: Can we remove this cast?
			ret.data = *reinterpret_cast<const T*>(GetCurrentAddress());
			ret.setOffset(m_Offset);
			ret.setSize(sizeof T);
			m_Offset += sizeof T;
			return ret;
		}

		// TODO: Do we need block versions of the string getters?
		std::string GetStringA(size_t cchChar = -1);
		std::wstring GetStringW(size_t cchChar = -1);
		std::vector<BYTE> GetBYTES(size_t cbBytes, size_t cbMaxBytes = -1);
		blockBytes GetBlockBYTES(size_t cbBytes, size_t cbMaxBytes = -1)
		{
			auto ret = blockBytes();
			ret.setOffset(m_Offset);
			ret.data = GetBYTES(cbBytes, cbMaxBytes);
			ret.setSize(ret.data.size() * sizeof(BYTE));
			return ret;
		}

		std::vector<BYTE> GetRemainingData();

	private:
		bool CheckRemainingBytes(size_t cbBytes) const;
		std::vector<BYTE> m_Bin;
		size_t m_Offset{};
	};
}