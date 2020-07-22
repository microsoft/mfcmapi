#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	class CMultiValuePropertyEditor : public CEditor
	{
	public:
		CMultiValuePropertyEditor(
			_In_ CWnd* pParentWnd,
			UINT uidTitle,
			bool bIsAB,
			_In_opt_ LPVOID lpAllocParent,
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
		void WriteSPropValueToObject() const;
		void WriteMultiValueStringsToSPropValue();
		void UpdateListRow(_In_ LPSPropValue lpProp, ULONG iMVCount) const;
		std::vector<LONG> GetLongArray() const;
		std::vector<std::vector<BYTE>> GetBinaryArray() const;
		void UpdateSmartView() const;
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void OnOK() override;

		// source variables
		LPMAPIPROP m_lpMAPIProp{};
		ULONG m_ulPropTag{};
		bool m_bIsAB{}; // whether the tag is from the AB or not
		const _SPropValue* m_lpsInputValue{};
		LPSPropValue m_lpsOutputValue{};

		// all calls to MAPIAllocateMore will use m_lpAllocParent
		// this is not something to be freed
		LPVOID m_lpAllocParent{};
	};
} // namespace dialog::editor