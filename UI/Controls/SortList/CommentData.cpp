#include <StdAfx.h>
#include <UI/Controls/SortList/CommentData.h>

namespace controls
{
	namespace sortlistdata
	{
		CommentData::CommentData(_In_ const _SPropValue* lpOldProp)
		{
			m_lpOldProp = lpOldProp;
			m_lpNewProp = nullptr;
		}
	}
}