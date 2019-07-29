#include <StdAfx.h>
#include <UI/Controls/SortList/ResData.h>

namespace controls
{
	namespace sortlistdata
	{
		ResData::ResData(_In_opt_ const _SRestriction* lpOldRes)
		{
			m_lpOldRes = lpOldRes;
			m_lpNewRes = nullptr;
		}
	} // namespace sortlistdata
} // namespace controls