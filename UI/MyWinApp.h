#pragma once
// Main header file for the MFCMAPI application

#ifndef __AFXWIN_H__
#error include 'stdafx.h' before including this file for PCH
#endif

namespace ui
{
	class CMyWinApp : public CWinApp
	{
	public:
		CMyWinApp();
	private:
		BOOL InitInstance() override;
	};
}