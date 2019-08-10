#pragma once
#include <core/sortlistdata/data.h>

namespace controls
{
	namespace sortlistdata
	{
		class propListData : public IData
		{
		public:
			propListData(_In_ ULONG ulPropTag);
			ULONG m_ulPropTag;
		};
	} // namespace sortlistdata
} // namespace controls