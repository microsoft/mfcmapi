#pragma once

namespace controls
{
	class CScrollHost : public CWnd
	{
	public:
		_Check_return_ BOOL Create(_In_ CWnd* parent);
		_Check_return_ int GetTrackPosition() const noexcept { return m_trackPosition; }

	private:
		struct ScrollGeometry
		{
			RECT bounds{};
			int trackTop{};
			int trackHeight{};
			int thumbTop{};
			int thumbHeight{};
			int minimum{};
			int maximumPosition{};
		};

		_Check_return_ bool GetScrollGeometry(_Out_ ScrollGeometry* geometry) const noexcept;
		void StopTracking() noexcept;
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;

		HWND m_parent{};
		int m_dragOffset{-1};
		int m_repeatCommand{SB_ENDSCROLL};
		int m_trackPosition{};
	};
} // namespace controls
