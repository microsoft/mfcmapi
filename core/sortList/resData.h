#pragma once
#include <core/sortList/data.h>

namespace controls
{
	namespace sortlistdata
	{
		class resData : public IData
		{
		public:
			resData(_In_opt_ const _SRestriction* lpOldRes);
			const _SRestriction* m_lpOldRes; // not allocated - just a pointer
			LPSRestriction m_lpNewRes; // Owned by an alloc parent - do not free
		};
	} // namespace sortlistdata
} // namespace controls