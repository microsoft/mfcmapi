#include <core/stdafx.h>
#include <core/propertyBag/mapiPropPropertyBag.h>
#include <core/sortlistdata/sortListData.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/error.h>
#include <core/mapi/mapiMemory.h>
#include <core/model/mapiRowModel.h>

namespace propertybag
{
	mapiPropPropertyBag::mapiPropPropertyBag(
		_In_ LPMAPIPROP lpProp,
		_In_ sortlistdata::sortListData* lpListData,
		_In_ bool bIsAB)
		: m_lpProp(lpProp), m_lpListData(lpListData), m_bIsAB(bIsAB)
	{
		if (m_lpProp) m_lpProp->AddRef();

		if (mapi::IsABObject(m_lpProp)) m_bIsAB = true;
	}

	mapiPropPropertyBag::~mapiPropPropertyBag()
	{
		if (m_lpProp) m_lpProp->Release();
	}

	propBagFlags mapiPropPropertyBag::GetFlags() const
	{
		auto ulFlags = propBagFlags::None | propBagFlags::CanDelete;
		if (m_bIsAB) ulFlags |= propBagFlags::AB;
		if (m_bGetPropsSucceeded) ulFlags |= propBagFlags::BackedByGetProps;
		return ulFlags;
	}

	bool mapiPropPropertyBag::IsEqual(_In_ const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const
	{
		if (!lpPropBag) return false;
		if (GetType() != lpPropBag->GetType()) return false;

		const auto lpOther = std::dynamic_pointer_cast<mapiPropPropertyBag>(lpPropBag);
		if (lpOther)
		{
			if (m_lpListData != lpOther->m_lpListData) return false;
			if (m_lpProp != lpOther->m_lpProp) return false;
			return true;
		}

		return false;
	}

	_Check_return_ HRESULT mapiPropPropertyBag::Commit()
	{
		if (nullptr == m_lpProp) return S_OK;

		return WC_H(m_lpProp->SaveChanges(KEEP_OPEN_READWRITE));
	}

	// Always returns a propval, even in errors, unless we fail allocating memory
	_Check_return_ LPSPropValue mapiPropPropertyBag::GetOneProp(_In_ ULONG ulPropTag, const std::wstring& /*name*/)
	{
		auto lpPropRet = LPSPropValue{};
		auto hRes = WC_MAPI(mapi::HrGetOnePropEx(m_lpProp, ulPropTag, fMapiUnicode, &lpPropRet));

		// Special case for profile sections and row properties - we may have a property which was in our row that isn't available on the object
		// In that case, we'll get MAPI_E_NOT_FOUND, but the property will be present in m_lpListData->lpSourceProps
		// So we fetch it from there instead
		// The caller will assume the memory was allocated from them, so copy before handing it back
		if (hRes == MAPI_E_NOT_FOUND && m_lpListData)
		{
			const auto lpProp = m_lpListData->GetOneProp(ulPropTag);
			if (lpProp)
			{
				hRes = WC_H(mapi::DupeProps(1, lpProp, MAPIAllocateBuffer, &lpPropRet));
			}
		}

		// If we still don't have a prop, build an error prop
		if (!lpPropRet)
		{
			lpPropRet = mapi::allocate<LPSPropValue>(sizeof(SPropValue));
			if (lpPropRet)
			{
				lpPropRet->ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_ERROR);
				lpPropRet->Value.err = (hRes == S_OK) ? MAPI_E_NOT_FOUND : hRes;
			}
		}

		return lpPropRet;
	}

	void mapiPropPropertyBag::FreeBuffer(_In_ LPSPropValue lpProp)
	{
		// m_lpListData->lpSourceProps is the only data we might hand out that we didn't allocate
		// Don't delete it!!!
		if (m_lpListData && m_lpListData->getRow().lpProps == lpProp) return;

		if (lpProp) MAPIFreeBuffer(lpProp);
	}

	_Check_return_ HRESULT mapiPropPropertyBag::SetProps(_In_ ULONG cValues, _In_ LPSPropValue lpPropArray)
	{
		if (nullptr == m_lpProp) return S_OK;

		LPSPropProblemArray lpProblems = nullptr;
		const auto hRes = WC_H(m_lpProp->SetProps(cValues, lpPropArray, &lpProblems));
		EC_PROBLEMARRAY(lpProblems);
		MAPIFreeBuffer(lpProblems);
		return hRes;
	}

	_Check_return_ HRESULT
	mapiPropPropertyBag::SetProp(_In_ LPSPropValue lpProp, _In_ ULONG /*ulPropTag*/, const std::wstring& /*name*/)
	{
		if (nullptr == m_lpProp) return S_OK;

		return WC_H(HrSetOneProp(m_lpProp, lpProp));
	}

	_Check_return_ HRESULT mapiPropPropertyBag::DeleteProp(_In_ ULONG ulPropTag)
	{
		if (nullptr == m_lpProp) return S_OK;

		const auto hRes = mapi::DeleteProperty(m_lpProp, ulPropTag);
		if (hRes == MAPI_E_NOT_FOUND && PROP_TYPE(ulPropTag) == PT_ERROR)
		{
			// We failed to delete a property without giving a type.
			// If the original type was error, it was quite likely a stream property.
			// Let's guess some common stream types.
			if (SUCCEEDED(mapi::DeleteProperty(m_lpProp, CHANGE_PROP_TYPE(ulPropTag, PT_BINARY)))) return S_OK;
			if (SUCCEEDED(mapi::DeleteProperty(m_lpProp, CHANGE_PROP_TYPE(ulPropTag, PT_UNICODE)))) return S_OK;
			if (SUCCEEDED(mapi::DeleteProperty(m_lpProp, CHANGE_PROP_TYPE(ulPropTag, PT_STRING8)))) return S_OK;
		}

		return hRes;
	};

	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> mapiPropPropertyBag::GetAllModels()
	{
		if (nullptr == m_lpProp) return {};
		auto hRes = S_OK;
		m_bGetPropsSucceeded = false;

		if (!registry::useRowDataForSinglePropList)
		{
			const auto unicodeFlag = registry::preferUnicodeProps ? MAPI_UNICODE : fMapiUnicode;

			ULONG cValues{};
			LPSPropValue lpPropArray{};

			hRes = mapi::GetPropsNULL(m_lpProp, unicodeFlag, &cValues, &lpPropArray);
			if (SUCCEEDED(hRes))
			{
				m_bGetPropsSucceeded = true;
			}
			if (hRes == MAPI_E_CALL_FAILED)
			{
				// Some stores, like public folders, don't support properties on the root folder
				output::DebugPrint(output::dbgLevel::Generic, L"Failed to get call GetProps on this object!\n");
			}
			else if (FAILED(hRes)) // only report errors, not warnings
			{
				CHECKHRESMSG(hRes, IDS_GETPROPSNULLFAILED);
			}

			const auto models = model::propsToModels(cValues, lpPropArray, m_lpProp, m_bIsAB);
			MAPIFreeBuffer(lpPropArray);
			return models;
		}

		if (!m_bGetPropsSucceeded && m_lpListData)
		{
			auto models = std::vector<std::shared_ptr<model::mapiRowModel>>{};
			const auto row = m_lpListData->getRow();
			for (ULONG i = 0; i < row.cValues; i++)
			{
				const auto prop = row.lpProps[i];
				models.push_back(model::propToModel(&prop, prop.ulPropTag, m_lpProp, m_bIsAB));
			}

			return models;
		}

		return {};
	}
	_Check_return_ std::shared_ptr<model::mapiRowModel> mapiPropPropertyBag::GetOneModel(_In_ ULONG ulPropTag)
	{
		auto lpPropVal = LPSPropValue{};
		const auto hRes = WC_MAPI(mapi::HrGetOnePropEx(m_lpProp, ulPropTag, fMapiUnicode, &lpPropVal));
		if (SUCCEEDED(hRes) && lpPropVal)
		{
			const auto model = model::propToModel(lpPropVal, ulPropTag, m_lpProp, m_bIsAB);
			MAPIFreeBuffer(lpPropVal);
			return model;
		}

		if (lpPropVal) MAPIFreeBuffer(lpPropVal);
		lpPropVal = nullptr;

		// Special case for profile sections and row properties - we may have a property which was in our row that isn't available on the object
		// In that case, we'll get MAPI_E_NOT_FOUND, but the property will be present in m_lpListData->lpSourceProps
		// So we fetch it from there instead
		if (hRes == MAPI_E_NOT_FOUND && m_lpListData)
		{
			return model::propToModel(m_lpListData->GetOneProp(ulPropTag), ulPropTag, m_lpProp, m_bIsAB);
		}

		// If we still don't have a prop, build an error prop
		SPropValue propVal = {CHANGE_PROP_TYPE(ulPropTag, PT_ERROR), 0};
		propVal.Value.err = (hRes == S_OK) ? MAPI_E_NOT_FOUND : hRes;

		return model::propToModel(&propVal, ulPropTag, m_lpProp, m_bIsAB);
	}
} // namespace propertybag