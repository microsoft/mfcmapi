#include <core/stdafx.h>
#include <core/sortlistdata/resData.h>
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void resData::init(sortListData* data, _In_opt_ const _SRestriction* lpOldRes)
	{
		if (!data) return;

		data->init(std::make_shared<resData>(lpOldRes), true);
	}

	resData::resData(_In_opt_ const _SRestriction* lpOldRes) : m_lpOldRes(lpOldRes) {}
} // namespace sortlistdata