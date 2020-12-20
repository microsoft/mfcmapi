#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class resData : public IData
	{
	public:
		static void init(sortListData* data, _In_opt_ const _SRestriction* lpOldRes);

		resData(_In_opt_ const _SRestriction* lpOldRes) noexcept;

		_Check_return_ const _SRestriction* getCurrentRes()
		{
			if (m_lpNewRes)
			{
				return m_lpNewRes;
			}
			else
			{
				return m_lpOldRes;
			}
		}
		void setCurrentRes(_In_ _SRestriction* res) { m_lpNewRes = res; }
		_Check_return_ _SRestriction detachRes(_In_opt_ const VOID* parent);

	private:
		const _SRestriction* m_lpOldRes{}; // not allocated - just a pointer
		LPSRestriction m_lpNewRes{}; // Owned by an alloc parent - do not free
	};
} // namespace sortlistdata