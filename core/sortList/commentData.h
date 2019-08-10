#pragma once
#include <core/sortList/data.h>

namespace controls
{
	namespace sortlistdata
	{
		class commentData : public IData
		{
		public:
			commentData(_In_opt_ const _SPropValue* lpOldProp);
			const _SPropValue* m_lpOldProp; // not allocated - just a pointer
			LPSPropValue m_lpNewProp; // Owned by an alloc parent - do not free
		};
	} // namespace sortlistdata
} // namespace controls