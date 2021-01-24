#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class propModelData : public IData
	{
	public:
		static void init(sortListData* data, _In_ ULONG ulPropTag);

		propModelData(_In_ ULONG ulPropTag) noexcept;
		_Check_return_ ULONG getPropTag() { return m_ulPropTag; }

	private:
		ULONG m_ulPropTag{};
	};
} // namespace sortlistdata