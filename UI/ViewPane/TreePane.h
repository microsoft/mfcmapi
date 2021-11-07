#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/Controls/StyleTree/StyleTreeCtrl.h>

namespace viewpane
{
	class TreePane : public ViewPane
	{
	public:
		static std::shared_ptr<TreePane> Create(int paneID, UINT uidLabel, bool bReadOnly);

		std::function<void(controls::StyleTreeCtrl& tree)> InitializeCallback = nullptr;

		controls::StyleTreeCtrl m_Tree;

	private:
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		void CommitUIValues() override{};
		int GetFixedHeight() override;
		int GetLines() override { return m_bCollapsed ? 0 : 4; }
	};
} // namespace viewpane