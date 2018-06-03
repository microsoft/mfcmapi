#include <StdAfx.h>
#include <UI/Controls/SortList/ResData.h>

namespace controls
{
	namespace sortlistdata
	{
		ResData::ResData(_In_ const _SRestriction* lpOldRes)
		{
			m_lpOldRes = lpOldRes;
			m_lpNewRes = nullptr;
		}
	}
}