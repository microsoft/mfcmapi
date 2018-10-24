#pragma once
#include <UI/ViewPane/DropDownPane.h>
#include <UI/ViewPane/SplitterPane.h>
#include <UI/ViewPane/TreePane.h>
#include <Interpret/SmartView/SmartViewParser.h>

namespace viewpane
{
	class SmartViewPane : public DropDownPane
	{
	public:
		static SmartViewPane* Create(int paneID, UINT uidLabel);

		~SmartViewPane()
		{
			if (m_TreePane) m_TreePane->m_Tree.DeleteAllItems();
		}

		void SetStringW(const std::wstring& szMsg);
		void DisableDropDown();
		void SetParser(__ParsingTypeEnum iParser);
		void Parse(SBinary myBin);

	private:
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetFixedHeight() override;
		int GetLines() override;
		void RefreshTree(smartview::LPSMARTVIEWPARSER svp);
		void AddChildren(HTREEITEM parent, const smartview::block& data);
		void ItemSelected(HTREEITEM hItem);
		static void FreeNodeData(LPARAM lpData);

		void SetMargins(
			int iMargin,
			int iSideMargin,
			int iLabelHeight, // Height of the label
			int iSmallHeightMargin,
			int iLargeHeightMargin,
			int iButtonHeight, // Height of buttons below the control
			int iEditHeight) override; // height of an edit control

		SplitterPane m_Splitter;
		TreePane* m_TreePane{nullptr};
		bool m_bHasData{false};
		bool m_bDoDropDown{true};
	};
} // namespace viewpane