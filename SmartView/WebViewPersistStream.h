#pragma once
#include "SmartViewParser.h"

struct WebViewPersist
{
	DWORD dwVersion;
	DWORD dwType;
	DWORD dwFlags;
	vector<BYTE> dwUnused; // 7 DWORDs
	DWORD cbData;
	vector<BYTE> lpData;
};

class WebViewPersistStream : public SmartViewParser
{
public:
	WebViewPersistStream();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	DWORD m_cWebViews;
	vector<WebViewPersist> m_lpWebViews;
};