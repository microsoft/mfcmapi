#pragma once
#include <core/propertyBag/propertyBag.h>
#include <core/sortlistdata/sortListData.h>
#include <core/model/mapiRowModel.h>

namespace propertybag
{
	class rowPropertyBag : public IMAPIPropertyBag
	{
	public:
		rowPropertyBag(_In_ sortlistdata::sortListData* lpListData, _In_ bool bIsAB);
		virtual ~rowPropertyBag() = default;
		rowPropertyBag(_In_ const rowPropertyBag&) = delete;
		rowPropertyBag& operator=(const rowPropertyBag&) = delete;

		propBagType GetType() const override { return propBagType::Row; }
		bool IsEqual(_In_ const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const override;

		// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
		_Check_return_ LPMAPIPROP GetMAPIProp() const override { return nullptr; }

		_Check_return_ HRESULT Commit() override { return E_NOTIMPL; }
		_Check_return_ LPSPropValue GetOneProp(_In_ ULONG ulPropTag) override;
		// None of our GetProps allocate anything, so nothing to do here
		void FreeBuffer(LPSPropValue) override { return; }
		// TODO: This is for paste, something we don't yet support for rows
		_Check_return_ HRESULT SetProps(_In_ ULONG, _In_ LPSPropValue) override { return E_NOTIMPL; }
		_Check_return_ HRESULT SetProp(_In_ LPSPropValue lpProp) override;
		//TODO: Not supported yet
		_Check_return_ HRESULT DeleteProp(_In_ ULONG) override { return E_NOTIMPL; };
		_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> GetAllModels() override;
		_Check_return_ std::shared_ptr<model::mapiRowModel> GetOneModel(_In_ ULONG ulPropTag) override;

		_Check_return_ HRESULT GetAllProps(_In_ ULONG FAR* lpcValues, _In_ LPSPropValue FAR* lppPropArray);

	protected:
		propBagFlags GetFlags() const override;

	private:
		sortlistdata::sortListData* m_lpListData{};

		ULONG m_cValues{};
		LPSPropValue m_lpProps{};
		bool m_bRowModified{};
		bool m_bIsAB{};
		// For failure case in GetOneProp
		SPropValue m_missingProp{};
	};
} // namespace propertybag