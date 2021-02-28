#pragma once

namespace memory
{
	LPBYTE ByteVectorToLPBYTE(const std::vector<BYTE>& bin) noexcept;
	size_t align(size_t s) noexcept;
	template <typename T> void assign(std::vector<BYTE>& v, const size_t offset, const T val) noexcept
	{
		if (offset > v.size() + sizeof(T)) return;
		*(reinterpret_cast<T*>(&v[offset])) = val;
	}
} // namespace memory