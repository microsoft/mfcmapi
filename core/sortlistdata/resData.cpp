#include <core/stdafx.h>
#include <core/sortlistdata/resData.h>
#include <core/sortlistdata/sortListData.h>
#include <core/utility/error.h>

namespace sortlistdata
{
	void resData::init(sortListData* data, _In_opt_ const _SRestriction* lpOldRes)
	{
		if (!data) return;

		data->init(std::make_shared<resData>(lpOldRes), true);
	}

	resData::resData(_In_opt_ const _SRestriction* lpOldRes) noexcept : m_lpOldRes(lpOldRes) {}

	_Check_return_ _SRestriction resData::detachRes(_In_opt_ const VOID* parent)
	{
		auto ret = _SRestriction{};
		if (m_lpNewRes)
		{
			memcpy(&ret, m_lpNewRes, sizeof(SRestriction));
			// clear out members so we don't double free
			memset(m_lpNewRes, 0, sizeof(SRestriction));
		}
		else
		{
			EC_H_S(mapi::HrCopyRestrictionArray(m_lpOldRes, parent, 1, &ret));
		}

		return ret;
	}

} // namespace sortlistdata