#pragma once
#include <UI/ViewPane/ViewPane.h>

namespace viewpane
{
	class AddInApprovalPane final : public ViewPane
	{
	public:
		static std::shared_ptr<AddInApprovalPane> Create(int paneID);
		ULONG HandleChange(UINT nID) override;
		bool containsWindow(HWND hWnd) const noexcept override;
		RECT GetWindowRect() const noexcept override;

	private:
		void Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc) override;
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		void CommitUIValues() override {}
		int GetMinWidth() override;
		int GetFixedHeight() override;

		CButton m_Button;
		int m_iButtonWidth{};
	};
} // namespace viewpane
