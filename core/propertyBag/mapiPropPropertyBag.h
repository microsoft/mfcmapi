#pragma once
#include <core/propertyBag/propertyBag.h>
#include <core/sortlistdata/sortListData.h>
#include <core/model/mapiRowModel.h>

namespace propertybag
{
	class mapiPropPropertyBag : public IMAPIPropertyBag
	{
	public:
		mapiPropPropertyBag(LPMAPIPROP lpProp, sortlistdata::sortListData* lpListData, bool bIsAB);
		~mapiPropPropertyBag();
		mapiPropPropertyBag(const mapiPropPropertyBag&) = delete;
		mapiPropPropertyBag& operator=(const mapiPropPropertyBag&) = delete;

		propBagType GetType() const override { return propBagType::MAPIProp; }
		bool IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const override;

		// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
		_Check_return_ LPMAPIPROP GetMAPIProp() const override { return m_lpProp; }

		_Check_return_ HRESULT Commit() override;
		_Check_return_ LPSPropValue GetOneProp(ULONG ulPropTag) override;
		void FreeBuffer(LPSPropValue lpProp) override;
		_Check_return_ HRESULT SetProps(ULONG cValues, LPSPropValue lpPropArray) override;
		_Check_return_ HRESULT SetProp(LPSPropValue lpProp) override;
		_Check_return_ HRESULT DeleteProp(ULONG ulPropTag) override;
		_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> GetAllModels() override;
		_Check_return_ std::shared_ptr<model::mapiRowModel> GetOneModel(ULONG ulPropTag) override;

	protected:
		propBagFlags GetFlags() const override;

	private:
		sortlistdata::sortListData* m_lpListData{};
		LPMAPIPROP m_lpProp{};
		bool m_bGetPropsSucceeded{};
		bool m_bIsAB{};
	};
} // namespace propertybag