#include <core/stdafx.h>
#include <core/sortlistdata/propModelData.h>
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void propModelData::init(sortListData* data, _In_ ULONG ulPropTag)
	{
		if (!data) return;

		data->init(std::make_shared<propModelData>(ulPropTag), true);
	}

	propModelData::propModelData(_In_ ULONG ulPropTag) noexcept : m_ulPropTag(ulPropTag) {}
} // namespace sortlistdata