#pragma once
#include <UI/ViewPane/DropDownPane.h>
#include <UI/ViewPane/SplitterPane.h>
#include <UI/ViewPane/TreePane.h>
#include <core/smartview/block/block.h>

enum class parserType;

namespace viewpane
{
	class SmartViewPane : public DropDownPane
	{
	public:
		static std::shared_ptr<SmartViewPane> Create(int paneID, UINT uidLabel);

		SmartViewPane();
		void SetParser(parserType parser);
		void Parse(const std::vector<BYTE>& myBin) { Parse(std::vector<std::vector<BYTE>>{myBin}); }
		void Parse(const std::vector<std::vector<BYTE>>& myBins);
		std::function<void(smartview::block*)> OnItemSelected = nullptr;
		std::function<void(_In_ const SBinary& lpBin)> OnActionButton = nullptr;
		bool containsWindow(HWND hWnd) const noexcept override;

	private:
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetFixedHeight() override;
		int GetLines() override;
		ULONG HandleChange(UINT nID) override;
		void HandleAction();
		void AddChildren(HTREEITEM parent, const std::shared_ptr<smartview::block>& data);
		void ItemSelected(HTREEITEM hItem);
		void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* /*pResult*/, _In_ HTREEITEM hItemCurHover) const;
		void SetStringW(const std::wstring& szMsg);

		void SetMargins(
			int iMargin,
			int iSideMargin,
			int iLabelHeight, // Height of the label
			int iSmallHeightMargin,
			int iLargeHeightMargin,
			int iButtonHeight, // Height of buttons below the control
			int iEditHeight) override; // height of an edit control

		std::vector<std::vector<BYTE>> m_bins;
		std::shared_ptr<smartview::block> treeData = smartview::block::create();
		std::shared_ptr<SplitterPane> m_Splitter;
		std::shared_ptr<TreePane> m_TreePane;
		bool m_bHasData{false};
	};
} // namespace viewpane