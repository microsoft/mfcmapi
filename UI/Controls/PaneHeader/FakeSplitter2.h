#pragma once
namespace controls
{
	class CFakeSplitter2 : public CWnd
	{
	public:
		CFakeSplitter2() = default;
		~CFakeSplitter2();

		void Init(HWND hWnd);

		void SetRightLabel(const std::wstring szLabel);
		void OnSize(UINT nType, int cx, int cy);
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height);

	private:
		void OnPaint();
		int OnCreate(LPCREATESTRUCT lpCreateStruct);
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;

		CEdit m_rightLabel; // Label on right of header
		int m_rightLabelWidth{0}; // The width of the string

		HWND m_hwndParent{};

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls