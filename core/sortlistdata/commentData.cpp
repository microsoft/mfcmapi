#include <core/stdafx.h>
#include <core/sortlistdata/commentData.h>

namespace controls
{
	namespace sortlistdata
	{
		commentData::commentData(_In_opt_ const _SPropValue* lpOldProp)
		{
			m_lpOldProp = lpOldProp;
			m_lpNewProp = nullptr;
		}
	} // namespace sortlistdata
} // namespace controls