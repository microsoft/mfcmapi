#include <StdAfx.h>
#include <UI/Dialogs/Editors/propeditor/ipropeditor.h>
#include <UI/Dialogs/Editors/propeditor/PropertyEditor.h>
#include <UI/Dialogs/Editors/propeditor/MultiValuePropertyEditor.h>
#include <core/utility/output.h>
#include <core/utility/error.h>

namespace dialog::editor
{
	void WriteSPropValueToObject(_In_opt_ LPMAPIPROP lpMAPIProp, _In_opt_ const _SPropValue* lpNewValue)
	{
		if (!lpMAPIProp || !lpNewValue || !lpNewValue->ulPropTag) return;

		LPSPropProblemArray lpProblemArray = nullptr;

		const auto hRes = EC_MAPI(lpMAPIProp->SetProps(1, const_cast<LPSPropValue>(lpNewValue), &lpProblemArray));

		EC_PROBLEMARRAY(lpProblemArray);
		MAPIFreeBuffer(lpProblemArray);

		if (SUCCEEDED(hRes))
		{
			EC_MAPI_S(lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
		}
	}

	_Check_return_ std::shared_ptr<IPropEditor> DisplayPropertyEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		UINT uidTitle,
		const std::wstring& name,
		bool bIsAB,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		bool bMVRow,
		_In_opt_ const _SPropValue* lpsPropValue)
	{
		_SPropValue* sourceProp = nullptr;
		// We got a MAPI prop object and no input value, go look one up
		if (lpMAPIProp && !lpsPropValue)
		{
			auto sTag = SPropTagArray{
				1, PROP_TYPE(ulPropTag) == PT_ERROR ? CHANGE_PROP_TYPE(ulPropTag, PT_UNSPECIFIED) : ulPropTag};
			ULONG ulValues = NULL;

			WC_MAPI_S(lpMAPIProp->GetProps(&sTag, NULL, &ulValues, &sourceProp));
			lpsPropValue = sourceProp;
		}

		if (lpsPropValue)
		{
			// Missing values can be created, but other errors must not be read as the requested type.
			if (PROP_TYPE(lpsPropValue->ulPropTag) == PT_ERROR && lpsPropValue->Value.err == MAPI_E_NOT_FOUND &&
				PROP_TYPE(ulPropTag) != PT_ERROR)
			{
				lpsPropValue = nullptr;
			}
			else
			{
				ulPropTag = lpsPropValue->ulPropTag;
			}
		}

		auto propertyEditor = std::shared_ptr<IPropEditor>{};

		// Check for the multivalue prop case
		if (PROP_TYPE(ulPropTag) & MV_FLAG)
		{
			propertyEditor = std::make_shared<CMultiValuePropertyEditor>(
				pParentWnd, lpMapiObjects, uidTitle, name, bIsAB, lpMAPIProp, ulPropTag, lpsPropValue);
		}
		// Or the single value prop case
		else
		{
			propertyEditor = std::make_shared<CPropertyEditor>(
				pParentWnd, lpMapiObjects, uidTitle, name, bIsAB, bMVRow, lpMAPIProp, ulPropTag, lpsPropValue);
		}

		if (propertyEditor)
		{
			if (propertyEditor->DisplayDialog())
			{
				if (!bMVRow) WriteSPropValueToObject(lpMAPIProp, propertyEditor->getValue());
			}
			else
			{
				propertyEditor = {};
			}
		}

		MAPIFreeBuffer(sourceProp);

		return propertyEditor;
	}
} // namespace dialog::editor