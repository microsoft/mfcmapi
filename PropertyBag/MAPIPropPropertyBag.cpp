#include "stdafx.h"
#include "..\stdafx.h"
#include "MAPIPropPropertyBag.h"
#include "..\MAPIFunctions.h"
#include "..\Error.h"
#include "..\InterpretProp.h"

MAPIPropPropertyBag::MAPIPropPropertyBag(LPMAPIPROP lpProp, SortListData* lpListData)
{
	m_lpListData = lpListData;
	m_lpProp = lpProp;
	if (m_lpProp) m_lpProp->AddRef();
	m_bGetPropsSucceeded = false;
}

MAPIPropPropertyBag::~MAPIPropPropertyBag()
{
	if (m_lpProp) m_lpProp->Release();
}

ULONG MAPIPropPropertyBag::GetFlags()
{
	ULONG ulFlags = pbNone;
	if (m_bGetPropsSucceeded) ulFlags |= pbBackedByGetProps;
	return ulFlags;
}

propBagType MAPIPropPropertyBag::GetType()
{
	return pbMAPIProp;
}

bool MAPIPropPropertyBag::IsEqual(LPMAPIPROPERTYBAG lpPropBag)
{
	if (!lpPropBag) return false;
	if (GetType() != lpPropBag->GetType()) return false;

	MAPIPropPropertyBag* lpOther = static_cast<MAPIPropPropertyBag*>(lpPropBag);
	if (lpOther)
	{
		if (m_lpListData != lpOther->m_lpListData) return false;
		if (m_lpProp != lpOther->m_lpProp) return false;
		return true;
	}

	return false;
}

// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
_Check_return_ LPMAPIPROP MAPIPropPropertyBag::GetMAPIProp()
{
	return m_lpProp;
}

_Check_return_ HRESULT MAPIPropPropertyBag::Commit()
{
	if (NULL == m_lpProp) return S_OK;

	HRESULT hRes = S_OK;
	WC_H(m_lpProp->SaveChanges(KEEP_OPEN_READWRITE));
	return hRes;
}

_Check_return_ HRESULT MAPIPropPropertyBag::GetAllProps(
	ULONG FAR* lpcValues,
	LPSPropValue FAR* lppPropArray)
{
	if (NULL == m_lpProp) return S_OK;
	HRESULT hRes = S_OK;
	m_bGetPropsSucceeded = false;

	if (!RegKeys[regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST].ulCurDWORD)
	{
		hRes = ::GetPropsNULL(m_lpProp,
			fMapiUnicode,
			lpcValues,
			lppPropArray);
		if (SUCCEEDED(hRes))
		{
			m_bGetPropsSucceeded = true;
		}
		if (MAPI_E_CALL_FAILED == hRes)
		{
			// Some stores, like public folders, don't support properties on the root folder
			DebugPrint(DBGGeneric, L"Failed to get call GetProps on this object!\n");
		}
		else if (FAILED(hRes)) // only report errors, not warnings
		{
			CHECKHRESMSG(hRes, IDS_GETPROPSNULLFAILED);
		}
		return hRes;
	}

	if (!m_bGetPropsSucceeded && m_lpListData)
	{
		*lpcValues = m_lpListData->cSourceProps;
		*lppPropArray = m_lpListData->lpSourceProps;
		hRes = S_OK;
	}

	return hRes;
}

_Check_return_ HRESULT MAPIPropPropertyBag::GetProps(
	LPSPropTagArray lpPropTagArray,
	ULONG ulFlags,
	ULONG FAR* lpcValues,
	LPSPropValue FAR* lppPropArray)
{
	if (NULL == m_lpProp) return S_OK;

	HRESULT hRes = S_OK;
	WC_H(m_lpProp->GetProps(lpPropTagArray, ulFlags, lpcValues, lppPropArray));
	return hRes;
}

_Check_return_ HRESULT MAPIPropPropertyBag::GetProp(
	ULONG ulPropTag,
	LPSPropValue FAR* lppProp)
{
	if (NULL == m_lpProp) return S_OK;

	HRESULT hRes = S_OK;
	WC_MAPI(HrGetOneProp(m_lpProp, ulPropTag, lppProp));

	// Special case for profile sections and row properties - we may have a property which was in our row that isn't available on the object
	// In that case, we'll get MAPI_E_NOT_FOUND, but the property will be present in m_lpListData->lpSourceProps
	// So we fetch it from there instead
	// The caller will assume the memory was allocated from them, so copy before handing it back
	if (MAPI_E_NOT_FOUND == hRes && m_lpListData)
	{
		LPSPropValue lpProp = PpropFindProp(m_lpListData->lpSourceProps, m_lpListData->cSourceProps, ulPropTag);
		if (lpProp)
		{
			hRes = S_OK;
			WC_MAPI(ScDupPropset(1, lpProp, MAPIAllocateBuffer, lppProp));
		}
	}
	return hRes;
}

void MAPIPropPropertyBag::FreeBuffer(LPSPropValue lpProp)
{
	// m_lpListData->lpSourceProps is the only data we might hand out that we didn't allocate
	// Don't delete it!!!
	if (m_lpListData && m_lpListData->lpSourceProps == lpProp) return;

	if (lpProp) MAPIFreeBuffer(lpProp);
	return;
}

_Check_return_ HRESULT MAPIPropPropertyBag::SetProps(
	ULONG cValues,
	LPSPropValue lpPropArray)
{
	if (NULL == m_lpProp) return S_OK;

	HRESULT hRes = S_OK;
	LPSPropProblemArray lpProblems = NULL;
	WC_H(m_lpProp->SetProps(cValues, lpPropArray, &lpProblems));
	EC_PROBLEMARRAY(lpProblems);
	MAPIFreeBuffer(lpProblems);
	return hRes;
}

_Check_return_ HRESULT MAPIPropPropertyBag::SetProp(
	LPSPropValue lpProp)
{
	if (NULL == m_lpProp) return S_OK;

	HRESULT hRes = S_OK;
	WC_H(HrSetOneProp(m_lpProp, lpProp));
	return hRes;
}

_Check_return_ HRESULT MAPIPropPropertyBag::DeleteProp(
	ULONG ulPropTag)
{
	if (NULL == m_lpProp) return S_OK;

	// TODO: eliminate DeleteProperty - not used much and might all could go through here.
	return DeleteProperty(m_lpProp, ulPropTag);
};