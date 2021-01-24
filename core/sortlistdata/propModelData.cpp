#include <core/stdafx.h>
#include <core/sortlistdata/propModelData.h>
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void propModelData::init(sortListData* data, _In_ std::shared_ptr<model::mapiRowModel> model)
	{
		if (!data) return;

		data->init(std::make_shared<propModelData>(model), true);
	}

	propModelData::propModelData(_In_ std::shared_ptr<model::mapiRowModel> model) noexcept : m_model(model) {}
} // namespace sortlistdata