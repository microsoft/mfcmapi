#include <core/stdafx.h>
#include <core/utility/memory.h>

namespace memory
{
	// Converts vector<BYTE> to LPBYTE allocated with new
	LPBYTE ByteVectorToLPBYTE(const std::vector<BYTE>& bin) noexcept
	{
		if (bin.empty()) return nullptr;

		const auto lpBin = new (std::nothrow) BYTE[bin.size()];
		if (lpBin != nullptr)
		{
			memset(lpBin, 0, bin.size());
			memcpy(lpBin, &bin[0], bin.size());
		}

		return lpBin;
	}

	size_t align(size_t s) noexcept
	{
		constexpr auto _align = static_cast<size_t>(sizeof(DWORD) - 1);
		if (s == ULONG_MAX || (s + _align) < s) return ULONG_MAX;

		return (((s) + _align) & ~_align);
	}
} // namespace memory