#pragma once

// Enable this to skip double buffering while investigating drawing issues
//#define SKIPBUFFER

class CDoubleBuffer
{
private:
	HDC m_hdcMem;
	HBITMAP m_hbmpMem;
	HDC m_hdcPaint;
	RECT m_rcPaint;

	void Cleanup();

public:
	CDoubleBuffer();
	~CDoubleBuffer();
	void Begin(_Inout_ HDC& hdc, _In_ const RECT& prcPaint);
	void End(_Inout_ HDC& hdc);
};
