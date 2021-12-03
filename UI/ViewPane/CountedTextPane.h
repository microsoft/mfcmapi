#pragma once
#include <UI/ViewPane/TextPane.h>

namespace viewpane
{
	class CountedTextPane : public TextPane
	{
	public:
		static std::shared_ptr<CountedTextPane> Create(int paneID, UINT uidLabel, bool bReadOnly, UINT uidCountLabel);
		void SetCount(size_t iCount);
		bool containsWindow(HWND hWnd) const noexcept override;

	private:
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetFixedHeight() override;
		int GetLines() override;

		std::wstring m_szCountLabel; // String name of the count
		size_t m_iCount{0}; // The numeric count
	};
} // namespace viewpane