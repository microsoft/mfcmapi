#pragma once
#include <UI/Dialogs/Editors/Editor.h>
#include <core/mapi/cache/mapiObjects.h>

namespace dialog::editor
{
	class IPropEditor : public CEditor
	{
	public:
		IPropEditor(_In_opt_ CWnd* pParentWnd, UINT uidTitle, UINT uidPrompt, ULONG ulButtonFlags)
			: CEditor(pParentWnd, uidTitle, uidPrompt, ulButtonFlags)
		{
		}
		// Implementations of getValue returns data owned completly by the object
		// so there is nothing to free
		// Callers CANNOT hold on to this data
		// If they need it, they can use mapi::HrDupPropset
		virtual _Check_return_ LPSPropValue getValue() = 0;
	};

	// If lpAllocParent is passed, our SPropValue will be allocated off of it
	// Otherwise caller will need to ensure the SPropValue is properly freed
	_Check_return_ std::shared_ptr<IPropEditor> DisplayPropertyEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		UINT uidTitle,
		const std::wstring& name,
		bool bIsAB,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		bool bMVRow,
		_In_opt_ const _SPropValue* lpsPropValue);
} // namespace dialog::editor