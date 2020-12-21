#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class commentData : public IData
	{
	public:
		commentData(_In_opt_ const _SPropValue* lpOldProp) noexcept;

		static void init(sortListData* data, _In_opt_ const _SPropValue* lpOldProp);

		_Check_return_ const SPropValue* getCurrentProp() { return m_lpNewProp ? m_lpNewProp : m_lpOldProp; }
		void setCurrentProp(_In_ SPropValue* prop) { m_lpNewProp = prop; }

	private:
		const SPropValue* m_lpOldProp; // not allocated - just a pointer
		LPSPropValue m_lpNewProp; // Owned by an alloc parent - do not free
	};
} // namespace sortlistdata