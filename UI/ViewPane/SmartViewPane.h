#pragma once
#include <UI/ViewPane/DropDownPane.h>
#include <UI/ViewPane/TextPane.h>

namespace viewpane
{
	class SmartViewPane : public DropDownPane
	{
	public:
		static SmartViewPane* Create(UINT uidLabel);

		void SetStringW(const std::wstring& szMsg);
		void DisableDropDown();
		void SetParser(__ParsingTypeEnum iParser);
		void Parse(SBinary myBin);

	private:
		SmartViewPane();
		void Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc) override;
		void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetFixedHeight() override;
		int GetLines() override;

		void SetMargins(
			int iMargin,
			int iSideMargin,
			int iLabelHeight, // Height of the label
			int iSmallHeightMargin,
			int iLargeHeightMargin,
			int iButtonHeight, // Height of buttons below the control
			int iEditHeight) override; // height of an edit control

		TextPane m_TextPane;
		bool m_bHasData;
		bool m_bDoDropDown;
	};
} // namespace viewpane