#pragma once
#include <UI/Controls/SortList/Data.h>

namespace controls
{
	namespace sortlistdata
	{
		class CommentData :public IData
		{
		public:
			CommentData(_In_ const _SPropValue* lpOldProp);
			const _SPropValue* m_lpOldProp; // not allocated - just a pointer
			LPSPropValue m_lpNewProp; // Owned by an alloc parent - do not free
		};
	}
}