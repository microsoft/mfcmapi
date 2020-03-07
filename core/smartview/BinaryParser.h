#pragma once

namespace smartview
{
	// binaryParser - helper class for parsing binary data without
	// worrying about whether you've run off the end of your buffer.
	class binaryParser
	{
	public:
		binaryParser() = default;
		binaryParser(size_t cb, _In_count_(cb) const BYTE* _bin)
		{
			bin = _bin && cb ? std::vector<BYTE>(_bin, _bin + cb) : std::vector<BYTE>{};
			size = bin.size();
		}
		binaryParser(const binaryParser&) = delete;
		binaryParser& operator=(const binaryParser&) = delete;

		bool empty() const noexcept { return offset == size; }
		void advance(size_t cb) noexcept { offset += cb; }
		void rewind() noexcept { offset = 0; }
		size_t getOffset() const { return offset; }
		void setOffset(size_t _offset) { offset = _offset; }
		const BYTE* getAddress() const { return bin.data() + offset; }
		void setCap(size_t cap)
		{
			sizes.push(size);
			if (cap != 0 && offset + cap < size)
			{
				size = offset + cap;
			}
		}

		void clearCap()
		{
			if (sizes.empty())
			{
				size = bin.size();
			}
			else
			{
				size = sizes.top();
				sizes.pop();
			}
		}

		// If we're before the end of the buffer, return the count of remaining bytes
		// If we're at or past the end of the buffer, return 0
		// If we're before the beginning of the buffer, return 0
		size_t getSize() const { return offset > size ? 0 : size - offset; }
		bool checkSize(size_t cb) const { return cb <= getSize(); }

	private:
		std::vector<BYTE> bin;
		size_t offset{};
		size_t size{}; // When uncapped, this is bin.size(). When capped, this is our artificial capped size.
		std::stack<size_t> sizes;
	};
} // namespace smartview