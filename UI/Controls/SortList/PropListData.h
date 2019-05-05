#pragma once
#include <UI/Controls/SortList/Data.h>

namespace controls
{
	namespace sortlistdata
	{
		class PropListData : public IData
		{
		public:
			PropListData(_In_ ULONG ulPropTag);
			ULONG m_ulPropTag;
		};
	} // namespace sortlistdata
} // namespace controls