#pragma once

namespace smartview
{
	class blockBytes;
	class block
	{
	public:
		block() : offset(0), cb(0), text(L""), header(true) {}
		block(const std::wstring& _text) : offset(0), cb(0), text(_text), header(true) {}

		std::wstring ToString() const
		{
			std::vector<std::wstring> items;
			items.emplace_back(text);
			for (const auto& item : children)
			{
				items.emplace_back(item.ToString());
			}

			return strings::join(items, strings::emptystring);
		}
		size_t getSize() const { return cb; }
		void setSize(size_t _size) { cb = _size; }
		size_t getOffset() const { return offset; }
		void setOffset(size_t _offset) { offset = _offset; }
		template <typename... Args> void addHeader(const std::wstring& _text, Args... args)
		{
			children.emplace_back(block(strings::formatmessage(_text.c_str(), args...)));
		}
		void addBlock(const block& child, const std::wstring& _text)
		{
			auto block = child;
			block.text = _text;
			children.push_back(block);
		}
		template <typename... Args> void addBlock(const block& child, const std::wstring& _text, Args... args)
		{
			auto block = child;
			block.text = strings::formatmessage(_text.c_str(), args...);
			children.push_back(block);
		}
		void addBlock(const block& child) { children.push_back(child); }
		void addBlockBytes(const blockBytes& child);
		void addLine() { addHeader(L"\r\n"); }
		bool hasData() const { return !text.empty() || !children.empty(); }

	protected:
		size_t offset;
		size_t cb;

	private:
		std::wstring text;
		std::vector<block> children;
		bool header;
	};

	class blockBytes : public block
	{
	public:
		void setData(const std::vector<BYTE>& data) { _data = data; }
		std::vector<BYTE> getData() const { return _data; }
		operator const std::vector<BYTE>&() const { return _data; }
		operator const std::vector<BYTE>() const { return _data; }
		size_t size() const noexcept { return _data.size(); }
		bool empty() const noexcept { return _data.empty(); }
		const BYTE* data() const noexcept { return _data.data(); }
		BYTE operator[](const size_t _Pos) { return _data[_Pos]; }

	private:
		std::vector<BYTE> _data;
	};

	template <class T> class blockT : public block
	{
	public:
		blockT() { memset(&data, 0, sizeof(data)); }
		void setData(const T& _data) { data = _data; }
		T getData() const { return data; }
		operator T&() { return data; }
		operator T() const { return data; }
		T& operator=(const blockT<T>& _data)
		{
			data = _data.data;
			cb = _data.getSize();
			offset = _data.getOffset();
			return *this;
		}
		T& operator=(const T& _data)
		{
			data = _data;
			return *this;
		}

	private:
		T data;
	};

	class blockStringA : public block
	{
	public:
		void setData(const std::string& _data) { data = _data; }
		std::string getData() const { return data; }
		operator std::string&() { return data; }
		operator std::string() const { return data; }
		_Ret_z_ const char* c_str() const noexcept { return data.c_str(); }
		size_t length() const noexcept { return data.length(); }
		bool empty() const noexcept { return data.empty(); }

	private:
		std::string data;
	};

	class blockStringW : public block
	{
	public:
		void setData(const std::wstring& _data) { data = _data; }
		std::wstring getData() const { return data; }
		operator std::wstring&() { return data; }
		operator std::wstring() const { return data; }
		_Ret_z_ const wchar_t* c_str() const noexcept { return data.c_str(); }
		size_t length() const noexcept { return data.length(); }
		bool empty() const noexcept { return data.empty(); }

	private:
		std::wstring data;
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
		bool CheckRemainingBytes(size_t cbBytes) const;
		std::vector<BYTE> m_Bin;
		size_t m_Offset{};
	};
} // namespace smartview