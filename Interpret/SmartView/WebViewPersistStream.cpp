#include <StdAfx.h>
#include <Interpret/SmartView/WebViewPersistStream.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
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
			const auto cbData = m_Parser.Get<DWORD>();

			// Must have at least cbData bytes left to be a valid flag
			if (m_Parser.RemainingBytes() < cbData) break;

			m_Parser.Advance(cbData);
			m_cWebViews++;
		}

		// Now we parse for real
		m_Parser.Rewind();

		const auto cWebViews = m_cWebViews;
		if (cWebViews && cWebViews < _MaxEntriesSmall)
		{
			for (ULONG i = 0; i < cWebViews; i++)
			{
				WebViewPersist webViewPersist;
				webViewPersist.dwVersion = m_Parser.Get<DWORD>();
				webViewPersist.dwType = m_Parser.Get<DWORD>();
				webViewPersist.dwFlags = m_Parser.Get<DWORD>();
				webViewPersist.dwUnused = m_Parser.GetBYTES(7 * sizeof(DWORD));
				webViewPersist.cbData = m_Parser.Get<DWORD>();
				webViewPersist.lpData = m_Parser.GetBYTES(webViewPersist.cbData, _MaxBytes);
				m_lpWebViews.push_back(webViewPersist);
			}
		}
	}

	_Check_return_ std::wstring WebViewPersistStream::ToStringInternal()
	{
		auto szWebViewPersistStream = strings::formatmessage(IDS_WEBVIEWSTREAMHEADER, m_cWebViews);
		for (ULONG i = 0; i < m_lpWebViews.size(); i++)
		{
			auto szVersion = interpretprop::InterpretFlags(flagWebViewVersion, m_lpWebViews[i].dwVersion);
			auto szType = interpretprop::InterpretFlags(flagWebViewType, m_lpWebViews[i].dwType);
			auto szFlags = interpretprop::InterpretFlags(flagWebViewFlags, m_lpWebViews[i].dwFlags);

			szWebViewPersistStream += strings::formatmessage(
				IDS_WEBVIEWHEADER,
				i,
				m_lpWebViews[i].dwVersion, szVersion.c_str(),
				m_lpWebViews[i].dwType, szType.c_str(),
				m_lpWebViews[i].dwFlags, szFlags.c_str());

			szWebViewPersistStream += strings::BinToHexString(m_lpWebViews[i].dwUnused, true);

			szWebViewPersistStream += strings::formatmessage(IDS_WEBVIEWCBDATA, m_lpWebViews[i].cbData);

			switch (m_lpWebViews[i].dwType)
			{
			case WEBVIEWURL:
			{
				szWebViewPersistStream += strings::formatmessage(IDS_WEBVIEWURL);
				szWebViewPersistStream += strings::BinToTextStringW(m_lpWebViews[i].lpData, false);
				break;
			}
			default:
				szWebViewPersistStream += strings::formatmessage(IDS_WEBVIEWDATA);
				szWebViewPersistStream += strings::BinToHexString(m_lpWebViews[i].lpData, true);
				break;
			}
		}

		return szWebViewPersistStream;
	}
}