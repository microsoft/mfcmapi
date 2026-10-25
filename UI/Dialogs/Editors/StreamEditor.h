#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	class CStreamEditor : public CEditor
	{
	public:
		CStreamEditor(
			_In_ CWnd* pParentWnd,
			UINT uidTitle,
			UINT uidPrompt,
			_In_ LPMAPIPROP lpMAPIProp,
			ULONG ulPropTag,
			bool bGuessType,
			bool bIsAB);
		~CStreamEditor();

		void DisableSave() noexcept;

	private:
		BOOL OnInitDialog() override;
		void OpenPropertyStream(bool bWrite);
		void ReadTextStreamFromProperty() const;
		void WriteTextStreamToProperty();
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void OnOK() override;
		void SetEditReadOnly(ULONG id) const;

		// source variables
		LPMAPIPROP m_lpMAPIProp;
		LPSTREAM m_lpStream;
		ULONG m_ulPropTag;
		bool m_bIsAB; // whether the tag is from the AB or not

		UINT m_iTextBox;
		UINT m_iBinBox;
		UINT m_iSmartViewBox;
		bool m_bDoSmartView;
		bool m_bDocFile;
		bool m_bAllowTypeGuessing;
		bool m_bDisableSave;
		ULONG m_ulEditorType;
		HRESULT m_StreamError;
	};
} // namespace dialog::editor