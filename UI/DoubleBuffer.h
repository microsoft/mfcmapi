#pragma once

namespace ui
{
	// Enable this to skip double buffering while investigating drawing issues
	//#define SKIPBUFFER

	class CDoubleBuffer
	{
	private:
		HDC m_hdcMem;
		HBITMAP m_hbmpMem;
		HDC m_hdcPaint;
		RECT m_rcPaint;

		void Cleanup() noexcept;

	public:
		CDoubleBuffer();
		~CDoubleBuffer();
		void Begin(_Inout_ HDC& hdc, _In_ const RECT& prcPaint) noexcept;
		void End(_Inout_ HDC& hdc) noexcept;
	};
} // namespace ui