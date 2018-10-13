#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/Controls/StyleTreeCtrl.h>

namespace viewpane
{
	class TreePane : public ViewPane
	{
	public:
		static TreePane* Create(int paneID, UINT uidLabel, bool bReadOnly);

		controls::StyleTreeCtrl m_Tree;
		std::function<void(TreePane& tree)> InitializeCallback = nullptr;

		HTREEITEM
		AddChildNode(
			_In_ const std::wstring& szName,
			HTREEITEM hParent,
			LPARAM lpData,
			const std::function<void(HTREEITEM hItem)>& callback)
		{
			return m_Tree.AddChildNode(szName, hParent, lpData, callback);
		};

	private:
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		void CommitUIValues() override{};
		int GetFixedHeight() override;
		int GetLines() override { return m_bCollapsed ? 0 : 4; }
	};
} // namespace viewpane