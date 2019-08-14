#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	void InitPropList(sortListData* data, _In_ ULONG ulPropTag);

	class propListData : public IData
	{
	public:
		propListData(_In_ ULONG ulPropTag);
		ULONG m_ulPropTag;
	};
} // namespace sortlistdata