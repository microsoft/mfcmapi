#pragma once
#include <core/propertyBag/propertyBag.h>
#include <core/sortlistdata/sortListData.h>
#include <core/model/mapiRowModel.h>

namespace propertybag
{
	class mapiPropPropertyBag : public IMAPIPropertyBag
	{
	public:
		mapiPropPropertyBag(_In_ LPMAPIPROP lpProp, _In_ sortlistdata::sortListData* lpListData, _In_ bool bIsAB);
		~mapiPropPropertyBag();
		mapiPropPropertyBag(_In_ const mapiPropPropertyBag&) = delete;
		mapiPropPropertyBag& operator=(const mapiPropPropertyBag&) = delete;

		propBagType GetType() const override { return propBagType::MAPIProp; }
		bool IsEqual(_In_ const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const override;

		// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
		_Check_return_ LPMAPIPROP GetMAPIProp() const override { return m_lpProp; }

		_Check_return_ HRESULT Commit() override;
		_Check_return_ LPSPropValue GetOneProp(_In_ ULONG ulPropTag) override;
		void FreeBuffer(LPSPropValue lpProp) override;
		_Check_return_ HRESULT SetProps(_In_ ULONG cValues, _In_ LPSPropValue lpPropArray) override;
		_Check_return_ HRESULT SetProp(_In_ LPSPropValue lpProp) override;
		_Check_return_ HRESULT DeleteProp(_In_ ULONG ulPropTag) override;
		_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> GetAllModels() override;
		_Check_return_ std::shared_ptr<model::mapiRowModel> GetOneModel(_In_ ULONG ulPropTag) override;

	protected:
		propBagFlags GetFlags() const override;

	private:
		sortlistdata::sortListData* m_lpListData{};
		LPMAPIPROP m_lpProp{};
		bool m_bGetPropsSucceeded{};
		bool m_bIsAB{};
	};
} // namespace propertybag