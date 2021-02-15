#pragma once
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/ViewPane/SmartViewPane.h>

namespace dialog::editor
{
	class IPropEditor : public CEditor
	{
	public:
		IPropEditor(_In_opt_ CWnd* pParentWnd, UINT uidTitle, UINT uidPrompt, ULONG ulButtonFlags)
			: CEditor(pParentWnd, uidTitle, uidPrompt, ulButtonFlags)
		{
		}
		virtual _Check_return_ LPSPropValue getValue() = 0;
	};

	// If lpAllocParent is passed, our SPropValue will be allocated off of it
	// Otherwise caller will need to ensure the SPropValue is properly freed
	_Check_return_ std::shared_ptr<IPropEditor> DisplayPropertyEditor(
		_In_ CWnd* pParentWnd,
		UINT uidTitle,
		bool bIsAB,
		_In_opt_ LPVOID lpAllocParent,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		ULONG ulPropTag,
		bool bMVRow,
		_In_opt_ const _SPropValue* lpsPropValue);

	class CPropertyEditor : public IPropEditor
	{
	public:
		CPropertyEditor(
			_In_ CWnd* pParentWnd,
			UINT uidTitle,
			bool bIsAB,
			bool bMVRow,
			_In_opt_ LPVOID lpAllocParent,
			_In_opt_ LPMAPIPROP lpMAPIProp,
			ULONG ulPropTag,
			_In_opt_ const _SPropValue* lpsPropValue);
		~CPropertyEditor();

		// Get values after we've done the DisplayDialog
		_Check_return_ LPSPropValue getValue() noexcept;

	private:
		BOOL OnInitDialog() override;
		void InitPropertyControls();
		void WriteStringsToSPropValue();
		void WriteSPropValueToObject() const;
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void OnOK() override;

		// source variables
		LPMAPIPROP m_lpMAPIProp{};
		ULONG m_ulPropTag{};
		bool m_bIsAB{}; // whether the tag is from the AB or not
		const _SPropValue* m_lpsInputValue{};
		LPSPropValue m_lpsOutputValue{};
		bool m_bDirty{};
		bool m_bMVRow{}; // whether this row came from a multivalued property. Used for smart view parsing.

		// all calls to MAPIAllocateMore will use m_lpAllocParent
		// this is not something to be freed
		LPVOID m_lpAllocParent{};
	};
} // namespace dialog::editor