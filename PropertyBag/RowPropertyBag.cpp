#include <StdAfx.h>
#include <PropertyBag/RowPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>

namespace propertybag
{
	RowPropertyBag::RowPropertyBag(controls::sortlistdata::SortListData* lpListData)
	{
		m_lpListData = lpListData;
		if (lpListData)
		{
			m_cValues = lpListData->cSourceProps;
			m_lpProps = lpListData->lpSourceProps;
		}
	}

	RowPropertyBag::~RowPropertyBag() = default;

	ULONG RowPropertyBag::GetFlags() const
	{
		ULONG ulFlags = pbNone;
		if (m_bRowModified) ulFlags |= pbModified;
		return ulFlags;
	}

	bool RowPropertyBag::IsEqual(const IMAPIPropertyBag* lpPropBag) const
	{
		if (!lpPropBag) return false;
		if (GetType() != lpPropBag->GetType()) return false;

		const auto lpOther = static_cast<const RowPropertyBag*>(lpPropBag);
		if (lpOther)
		{
			if (m_lpListData != lpOther->m_lpListData) return false;
			if (m_cValues != lpOther->m_cValues) return false;
			if (m_lpProps != lpOther->m_lpProps) return false;
			return true;
		}

		return false;
	}

	_Check_return_ HRESULT RowPropertyBag::Commit() { return E_NOTIMPL; }

	_Check_return_ HRESULT RowPropertyBag::GetAllProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray)
	{
		if (!lpcValues || !lppPropArray) return MAPI_E_INVALID_PARAMETER;

		// Just return what we have
		*lpcValues = m_cValues;
		*lppPropArray = m_lpProps;
		return S_OK;
	}

	_Check_return_ HRESULT RowPropertyBag::GetProps(
		LPSPropTagArray /*lpPropTagArray*/,
		ULONG /*ulFlags*/,
		ULONG FAR* /*lpcValues*/,
		LPSPropValue FAR* /*lppPropArray*/)
	{
		// This is only called from the Extra Props code. We can't support Extra Props from a row
		// so we don't need to implement this.
		return E_NOTIMPL;
	}

	_Check_return_ HRESULT RowPropertyBag::GetProp(ULONG ulPropTag, LPSPropValue FAR* lppProp)
	{
		if (!lppProp) return MAPI_E_INVALID_PARAMETER;

		*lppProp = PpropFindProp(m_lpProps, m_cValues, ulPropTag);
		return S_OK;
	}

	void RowPropertyBag::FreeBuffer(LPSPropValue /*lpProp*/)
	{
		// None of our GetProps allocate anything, so nothing to do here
		return;
	}

	// TODO: This is for paste, something we don't yet support for rows
	_Check_return_ HRESULT RowPropertyBag::SetProps(ULONG /*cValues*/, LPSPropValue /*lpPropArray*/)
	{
		return E_NOTIMPL;
	}

	// Concatenate two property arrays without duplicates
	// Entries in the first array trump entries in the second
	// Will also eliminate any duplicates already existing within the arrays
	_Check_return_ HRESULT ConcatLPSPropValue(
		ULONG ulVal1,
		_In_count_(ulVal1) LPSPropValue lpVal1,
		ULONG ulVal2,
		_In_count_(ulVal2) LPSPropValue lpVal2,
		_Out_ ULONG* lpulRetVal,
		_Deref_out_opt_ LPSPropValue* lppRetVal)
	{
		if (!lpulRetVal || !lppRetVal) return MAPI_E_INVALID_PARAMETER;
		if (ulVal1 && !lpVal1) return MAPI_E_INVALID_PARAMETER;
		if (ulVal2 && !lpVal2) return MAPI_E_INVALID_PARAMETER;
		*lpulRetVal = NULL;
		*lppRetVal = nullptr;
		auto hRes = S_OK;

		ULONG ulTargetArray = 0;
		ULONG ulNewArraySize = 0;
		LPSPropValue lpNewArray = nullptr;

		// Add the sizes of the passed in arrays
		if (ulVal2 && ulVal1)
		{
			ulNewArraySize = ulVal1;
			// Only count props in the second array if they're not in the first
			for (ULONG ulSourceArray = 0; ulSourceArray < ulVal2; ulSourceArray++)
			{
				if (!PpropFindProp(lpVal1, ulVal1, CHANGE_PROP_TYPE(lpVal2[ulSourceArray].ulPropTag, PT_UNSPECIFIED)))
				{
					ulNewArraySize++;
				}
			}
		}
		else
		{
			ulNewArraySize = ulVal1 + ulVal2;
		}

		if (ulNewArraySize)
		{
			// Allocate the base array - MyPropCopyMore will allocmore as needed for string/bin/etc
			lpNewArray = mapi::allocate<LPSPropValue>(ulNewArraySize * sizeof(SPropValue));
			if (lpNewArray)
			{
				if (ulVal1)
				{
					for (ULONG ulSourceArray = 0; ulSourceArray < ulVal1; ulSourceArray++)
					{
						if (!ulTargetArray || // if it's NULL, we haven't added anything yet
							!PpropFindProp(
								lpNewArray,
								ulTargetArray,
								CHANGE_PROP_TYPE(lpVal1[ulSourceArray].ulPropTag, PT_UNSPECIFIED)))
						{
							hRes = EC_H(mapi::MyPropCopyMore(
								&lpNewArray[ulTargetArray], &lpVal1[ulSourceArray], MAPIAllocateMore, lpNewArray));
							if (SUCCEEDED(hRes))
							{
								ulTargetArray++;
							}
							else
								break;
						}
					}
				}

				if (SUCCEEDED(hRes) && ulVal2)
				{
					for (ULONG ulSourceArray = 0; ulSourceArray < ulVal2; ulSourceArray++)
					{
						if (!ulTargetArray || // if it's NULL, we haven't added anything yet
							!PpropFindProp(
								lpNewArray,
								ulTargetArray,
								CHANGE_PROP_TYPE(lpVal2[ulSourceArray].ulPropTag, PT_UNSPECIFIED)))
						{
							// make sure we don't overrun.
							if (ulTargetArray >= ulNewArraySize)
							{
								hRes = MAPI_E_CALL_FAILED;
								break;
							}

							hRes = EC_H(mapi::MyPropCopyMore(
								&lpNewArray[ulTargetArray], &lpVal2[ulSourceArray], MAPIAllocateMore, lpNewArray));
							if (SUCCEEDED(hRes))
							{
								ulTargetArray++;
							}
							else
								break;
						}
					}
				}

				if (FAILED(hRes))
				{
					MAPIFreeBuffer(lpNewArray);
				}
				else
				{
					*lpulRetVal = ulTargetArray;
					*lppRetVal = lpNewArray;
				}
			}
		}

		return hRes;
	}

	_Check_return_ HRESULT RowPropertyBag::SetProp(LPSPropValue lpProp)
	{
		ULONG ulNewArray = NULL;
		LPSPropValue lpNewArray = nullptr;

		const auto hRes = EC_H(ConcatLPSPropValue(1, lpProp, m_cValues, m_lpProps, &ulNewArray, &lpNewArray));
		if (SUCCEEDED(hRes))
		{
			MAPIFreeBuffer(m_lpListData->lpSourceProps);
			m_lpListData->cSourceProps = ulNewArray;
			m_lpListData->lpSourceProps = lpNewArray;

			m_cValues = ulNewArray;
			m_lpProps = lpNewArray;
			m_bRowModified = true;
		}
		return hRes;
	}

	//TODO: Not supported yet
	_Check_return_ HRESULT RowPropertyBag::DeleteProp(ULONG /*ulPropTag*/) { return E_NOTIMPL; };
} // namespace propertybag
