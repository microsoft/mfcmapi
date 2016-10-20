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
	WebViewPersistStream();
	~WebViewPersistStream();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_cWebViews;
	WebViewPersistStruct* m_lpWebViews;
};