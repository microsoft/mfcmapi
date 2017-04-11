#include "StdAfx.h"
#include <UI/DoubleBuffer.h>

CDoubleBuffer::CDoubleBuffer() : m_hdcMem(nullptr), m_hbmpMem(nullptr), m_hdcPaint(nullptr)
{
	ZeroMemory(&m_rcPaint, sizeof m_rcPaint);
}

CDoubleBuffer::~CDoubleBuffer()
{
	Cleanup();
}

void CDoubleBuffer::Begin(_Inout_ HDC& hdc, _In_ const RECT& rcPaint)
{
#ifdef SKIPBUFFER
	UNREFERENCED_PARAMETER(hdc);
	UNREFERENCED_PARAMETER(rcPaint);
#else
	if (hdc)
	{
		m_hdcMem = CreateCompatibleDC(hdc);
		if (m_hdcMem)
		{
			m_hbmpMem = CreateCompatibleBitmap(
				hdc,
				rcPaint.right - rcPaint.left,
				rcPaint.bottom - rcPaint.top);

			if (m_hbmpMem)
			{
				(void)SelectObject(m_hdcMem, m_hbmpMem);
				(void)CopyRect(&m_rcPaint, &rcPaint);
				(void)OffsetWindowOrgEx(m_hdcMem,
					m_rcPaint.left,
					m_rcPaint.top,
					nullptr);

				(void)SelectObject(m_hdcMem, GetCurrentObject(hdc, OBJ_FONT));
				(void)SelectObject(m_hdcMem, GetCurrentObject(hdc, OBJ_BRUSH));
				(void)SelectObject(m_hdcMem, GetCurrentObject(hdc, OBJ_PEN));
				// cache the original DC and pass out the memory DC
				m_hdcPaint = hdc;
				hdc = m_hdcMem;
			}
			else
			{
				Cleanup();
			}
		}
	}
#endif
}

void CDoubleBuffer::End(_Inout_ HDC& hdc)
{
	if (hdc && hdc == m_hdcMem)
	{
		BitBlt(m_hdcPaint,
			m_rcPaint.left,
			m_rcPaint.top,
			m_rcPaint.right - m_rcPaint.left,
			m_rcPaint.bottom - m_rcPaint.top,
			m_hdcMem,
			m_rcPaint.left,
			m_rcPaint.top,
			SRCCOPY);

		// restore the original DC
		hdc = m_hdcPaint;

		Cleanup();
	}
}

void CDoubleBuffer::Cleanup()
{
	if (m_hbmpMem)
	{
		(void)DeleteObject(m_hbmpMem);
		m_hbmpMem = nullptr;
	}

	if (m_hdcMem)
	{
		(void)DeleteDC(m_hdcMem);
		m_hdcMem = nullptr;
	}

	m_hdcPaint = nullptr;
	ZeroMemory(&m_rcPaint, sizeof m_rcPaint);
}