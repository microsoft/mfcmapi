#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class propListData : public IData
	{
	public:
		static void init(sortListData* data, _In_ ULONG ulPropTag);

		ULONG m_ulPropTag{};

	private:
		propListData(_In_ ULONG ulPropTag);
	};
} // namespace sortlistdata