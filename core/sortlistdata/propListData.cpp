#include <core/stdafx.h>
#include <core/sortlistdata/propListData.h>
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void propListData::init(sortListData* data, _In_ ULONG ulPropTag)
	{
		if (!data) return;

		data->init(new (std::nothrow) propListData(ulPropTag), true);
	}

	propListData::propListData(_In_ ULONG ulPropTag) : m_ulPropTag(ulPropTag) {}
} // namespace sortlistdata