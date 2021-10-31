#pragma once
#include <UI/Dialogs/Editors/propeditor/ipropeditor.h>
#include <core/mapi/cache/mapiObjects.h>

namespace dialog::editor
{
	class CMultiValuePropertyEditor : public IPropEditor
	{
	public:
		CMultiValuePropertyEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			UINT uidTitle,
			const std::wstring& name,
			bool bIsAB,
			_In_opt_ LPMAPIPROP lpMAPIProp,
			ULONG ulPropTag,
			_In_opt_ const _SPropValue* lpsPropValue);
		~CMultiValuePropertyEditor();

		// Get values after we've done the DisplayDialog
		_Check_return_ LPSPropValue getValue() noexcept;
		_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData) override;

	private:
		BOOL OnInitDialog() override;
		void InitPropertyControls();
		void ReadMultiValueStringsFromProperty() const;
		void WriteMultiValueStringsToSPropValue();
		void UpdateListRow(_In_ LPSPropValue lpProp, ULONG iMVCount) const;
		std::vector<LONG> GetLongArray() const;
		std::vector<std::vector<BYTE>> GetBinaryArray() const;
		void UpdateSmartView() const;
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void OnOK() override;

		// source variables
		LPMAPIPROP m_lpMAPIProp{}; // Used only for parsing
		ULONG m_ulPropTag{};
		bool m_bIsAB{}; // whether the tag is from the AB or not
		const _SPropValue* m_lpsInputValue{};
		const std::wstring m_name;

		SPropValue m_sOutputValue{};
		std::vector<BYTE> m_bin; // Temp storage for m_sOutputValue
		std::vector<std::string> m_mvA; // Temp storage for m_sOutputValue array
		std::vector<std::wstring> m_mvW; // Temp storage for m_sOutputValue array
		std::vector<std::vector<BYTE>> m_mvBin; // Temp storage for m_sOutputValue array
		std::vector<GUID> m_mvGuid; // Temp storage for m_sOutputValue array
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};
} // namespace dialog::editor