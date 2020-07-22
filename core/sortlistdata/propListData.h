#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class propListData : public IData
	{
	public:
		static void init(sortListData* data, _In_ ULONG ulPropTag);

		propListData(_In_ ULONG ulPropTag) noexcept;

		ULONG m_ulPropTag{};
	};
} // namespace sortlistdata