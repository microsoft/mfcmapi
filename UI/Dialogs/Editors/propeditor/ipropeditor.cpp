#include <StdAfx.h>
#include <UI/Dialogs/Editors/propeditor/ipropeditor.h>
#include <UI/Dialogs/Editors/propeditor/PropertyEditor.h>
#include <UI/Dialogs/Editors/propeditor/MultiValuePropertyEditor.h>

namespace dialog::editor
{
	_Check_return_ std::shared_ptr<IPropEditor> DisplayPropertyEditor(
		_In_ CWnd* pParentWnd,
		UINT uidTitle,
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

			auto hRes = WC_MAPI(lpMAPIProp->GetProps(&sTag, NULL, &ulValues, &sourceProp));

			// Suppress MAPI_E_NOT_FOUND error when the source type is non error
			if (sourceProp && PROP_TYPE(sourceProp->ulPropTag) == PT_ERROR &&
				sourceProp->Value.err == MAPI_E_NOT_FOUND && PROP_TYPE(ulPropTag) != PT_ERROR)
			{
				MAPIFreeBuffer(sourceProp);
				sourceProp = nullptr;
			}

			if (hRes == MAPI_E_CALL_FAILED)
			{
				// Just suppress this - let the user edit anyway
				hRes = S_OK;
			}

			// In all cases where we got a value back, we need to reset our property tag to the value we got
			// This will address when the source is PT_UNSPECIFIED, when the returned value is PT_ERROR,
			// or any other case where the returned value has a different type than requested
			if (SUCCEEDED(hRes) && sourceProp)
			{
				ulPropTag = sourceProp->ulPropTag;
			}

			lpsPropValue = sourceProp;
		}
		else if (lpsPropValue && !ulPropTag)
		{
			ulPropTag = lpsPropValue->ulPropTag;
		}

		auto MyPropertyEditor = std::shared_ptr<IPropEditor>{};

		// Check for the multivalue prop case
		if (PROP_TYPE(ulPropTag) & MV_FLAG)
		{
			MyPropertyEditor = std::make_shared<CMultiValuePropertyEditor>(
				pParentWnd, uidTitle, bIsAB, lpMAPIProp, ulPropTag, lpsPropValue);
		}
		// Or the single value prop case
		else
		{
			MyPropertyEditor = std::make_shared<CPropertyEditor>(
				pParentWnd, uidTitle, bIsAB, bMVRow, lpMAPIProp, ulPropTag, lpsPropValue);
		}

		if (MyPropertyEditor && !MyPropertyEditor->DisplayDialog())
		{
			MyPropertyEditor = {};
		}

		MAPIFreeBuffer(sourceProp);

		return MyPropertyEditor;
	}
} // namespace dialog::editor