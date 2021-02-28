#pragma once
#include <core/sortlistdata/data.h>
#include <core/model/mapiRowModel.h>

namespace sortlistdata
{
	class sortListData;

	class propModelData : public IData
	{
	public:
		static void init(sortListData* data, _In_ std::shared_ptr<model::mapiRowModel> model);

		propModelData(_In_ std::shared_ptr<model::mapiRowModel> model) noexcept;
		_Check_return_ ULONG getPropTag() { return m_model->ulPropTag(); }
		_Check_return_ std::wstring getName() { return m_model->name(); }

	private:
		std::shared_ptr<model::mapiRowModel> m_model;
	};
} // namespace sortlistdata