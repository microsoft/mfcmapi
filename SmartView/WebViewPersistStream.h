#pragma once
#include "SmartViewParser.h"

struct WebViewPersistStruct
{
	DWORD dwVersion;
	DWORD dwType;
	DWORD dwFlags;
	DWORD dwUnused[7];
	DWORD cbData;
	LPBYTE lpData;
};

class WebViewPersistStream : public SmartViewParser
{
public:
	WebViewPersistStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~WebViewPersistStream();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	DWORD m_cWebViews;
	WebViewPersistStruct* m_lpWebViews;
};