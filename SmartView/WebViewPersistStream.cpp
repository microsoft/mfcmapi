#include "stdafx.h"
#include "WebViewPersistStream.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

WebViewPersistStream::WebViewPersistStream()
{
	m_cWebViews = 0;
}

void WebViewPersistStream::Parse()
{
	// Run through the parser once to count the number of web view structs
	for (;;)
	{
		// Must have at least 2 bytes left to have another struct
		if (m_Parser.RemainingBytes() < sizeof(DWORD) * 11) break;
		m_Parser.Advance(sizeof(DWORD) * 10);
		DWORD cbData;
		m_Parser.GetDWORD(&cbData);

		// Must have at least cbData bytes left to be a valid flag
		if (m_Parser.RemainingBytes() < cbData) break;

		m_Parser.Advance(cbData);
		m_cWebViews++;
	}

	// Now we parse for real
	m_Parser.Rewind();

	auto cWebViews = m_cWebViews;
	if (cWebViews && cWebViews < _MaxEntriesSmall)
	{
		for (ULONG i = 0; i < cWebViews; i++)
		{
			WebViewPersist webViewPersist;
			m_Parser.GetDWORD(&webViewPersist.dwVersion);
			m_Parser.GetDWORD(&webViewPersist.dwType);
			m_Parser.GetDWORD(&webViewPersist.dwFlags);
			webViewPersist.dwUnused = m_Parser.GetBYTES(7 * sizeof(DWORD));
			m_Parser.GetDWORD(&webViewPersist.cbData);
			webViewPersist.lpData = m_Parser.GetBYTES(webViewPersist.cbData, _MaxBytes);
			m_lpWebViews.push_back(webViewPersist);
		}
	}
}

_Check_return_ wstring WebViewPersistStream::ToStringInternal()
{
	auto szWebViewPersistStream = formatmessage(IDS_WEBVIEWSTREAMHEADER, m_cWebViews);
	for (ULONG i = 0; i < m_lpWebViews.size(); i++)
	{
		auto szVersion = InterpretFlags(flagWebViewVersion, m_lpWebViews[i].dwVersion);
		auto szType = InterpretFlags(flagWebViewType, m_lpWebViews[i].dwType);
		auto szFlags = InterpretFlags(flagWebViewFlags, m_lpWebViews[i].dwFlags);

		szWebViewPersistStream += formatmessage(
			IDS_WEBVIEWHEADER,
			i,
			m_lpWebViews[i].dwVersion, szVersion.c_str(),
			m_lpWebViews[i].dwType, szType.c_str(),
			m_lpWebViews[i].dwFlags, szFlags.c_str());

		szWebViewPersistStream += BinToHexString(m_lpWebViews[i].dwUnused, true);

		szWebViewPersistStream += formatmessage(IDS_WEBVIEWCBDATA, m_lpWebViews[i].cbData);

		switch (m_lpWebViews[i].dwType)
		{
		case WEBVIEWURL:
		{
			szWebViewPersistStream += formatmessage(IDS_WEBVIEWURL);
			szWebViewPersistStream += BinToTextStringW(m_lpWebViews[i].lpData, false);
			break;
		}
		default:
			szWebViewPersistStream += formatmessage(IDS_WEBVIEWDATA);
			szWebViewPersistStream += BinToHexString(m_lpWebViews[i].lpData, true);
			break;
		}
	}

	return szWebViewPersistStream;
}