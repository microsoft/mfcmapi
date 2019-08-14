#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class resData : public IData
	{
	public:
		static void init(sortListData* data, _In_opt_ const _SRestriction* lpOldRes);

		const _SRestriction* m_lpOldRes{}; // not allocated - just a pointer
		LPSRestriction m_lpNewRes{}; // Owned by an alloc parent - do not free

	private:
		resData(_In_opt_ const _SRestriction* lpOldRes);
	};
} // namespace sortlistdata