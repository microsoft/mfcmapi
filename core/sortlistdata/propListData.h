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
		_Check_return_ ULONG getPropTag() { return m_ulPropTag; }
		void setPropTag(ULONG ulPropTag) { m_ulPropTag = ulPropTag; }

	private:
		ULONG m_ulPropTag{};
	};
} // namespace sortlistdata