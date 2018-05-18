#pragma once
#include "SmartViewParser.h"

struct WebViewPersist
{
	DWORD dwVersion;
	DWORD dwType;
	DWORD dwFlags;
	std::vector<BYTE> dwUnused; // 7 DWORDs
	DWORD cbData;
	std::vector<BYTE> lpData;
};

class WebViewPersistStream : public SmartViewParser
{
public:
	WebViewPersistStream();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	DWORD m_cWebViews;
	std::vector<WebViewPersist> m_lpWebViews;
};