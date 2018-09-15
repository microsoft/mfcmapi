#pragma once
#include <UI/ViewPane/TextPane.h>

namespace viewpane
{
	class CountedTextPane : public TextPane
	{
	public:
		static CountedTextPane* CountedTextPane::Create(int paneID, UINT uidLabel, bool bReadOnly, UINT uidCountLabel);
		void SetCount(size_t iCount);

	private:
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetMinWidth(_In_ HDC hdc) override;
		int GetFixedHeight() override;
		int GetLines() override;

		CEdit m_Count; // The display of the count
		std::wstring m_szCountLabel; // String name of the count
		int m_iCountLabelWidth{0}; // The width of the string
		size_t m_iCount{0}; // The numeric count
	};
} // namespace viewpane