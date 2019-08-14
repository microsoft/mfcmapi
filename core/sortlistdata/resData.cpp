#include <core/stdafx.h>
#include <core/sortlistdata/resData.h>
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void InitRes(sortListData* data, _In_opt_ const _SRestriction* lpOldRes)
	{
		if (!data) return;
		data->Init(new (std::nothrow) resData(lpOldRes), true);
	}

	resData::resData(_In_opt_ const _SRestriction* lpOldRes) : m_lpOldRes(lpOldRes) {}
} // namespace sortlistdata