#include <core/stdafx.h>
#include <core/sortlistdata/resData.h>

namespace controls
{
	namespace sortlistdata
	{
		resData::resData(_In_opt_ const _SRestriction* lpOldRes)
		{
			m_lpOldRes = lpOldRes;
			m_lpNewRes = nullptr;
		}
	} // namespace sortlistdata
} // namespace controls