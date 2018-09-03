#include <StdAfx.h>
#include <Interpret/SmartView/WebViewPersistStream.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	WebViewPersistStream::WebViewPersistStream() { m_cWebViews = 0; }

	void WebViewPersistStream::Parse()
	{
		// Run through the parser once to count the number of web view structs
		for (;;)
		{
			// Must have at least 2 bytes left to have another struct
			if (m_Parser.RemainingBytes() < sizeof(DWORD) * 11) break;
			m_Parser.Advance(sizeof(DWORD) * 10);
			const auto cbData = m_Parser.GetBlock<DWORD>();

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
				webViewPersist.dwVersion = m_Parser.GetBlock<DWORD>();
				webViewPersist.dwType = m_Parser.GetBlock<DWORD>();
				webViewPersist.dwFlags = m_Parser.GetBlock<DWORD>();
				webViewPersist.dwUnused = m_Parser.GetBlockBYTES(7 * sizeof(DWORD));
				webViewPersist.cbData = m_Parser.GetBlock<DWORD>();
				webViewPersist.lpData = m_Parser.GetBlockBYTES(webViewPersist.cbData, _MaxBytes);
				m_lpWebViews.push_back(webViewPersist);
			}
		}
	}

	void WebViewPersistStream::ParseBlocks()
	{
		addHeader(L"Web View Persistence Object Stream\r\n");
		addHeader(L"cWebViews = %1!d!", m_cWebViews);
		for (ULONG i = 0; i < m_lpWebViews.size(); i++)
		{
			addLine();
			addLine();

			addHeader(L"Web View %1!d!\r\n", i);
			addBlock(
				m_lpWebViews[i].dwVersion,
				L"dwVersion = 0x%1!08X! = %2!ws!\r\n",
				m_lpWebViews[i].dwVersion.getData(),
				interpretprop::InterpretFlags(flagWebViewVersion, m_lpWebViews[i].dwVersion).c_str());
			addBlock(
				m_lpWebViews[i].dwType,
				L"dwType = 0x%1!08X! = %2!ws!\r\n",
				m_lpWebViews[i].dwType.getData(),
				interpretprop::InterpretFlags(flagWebViewType, m_lpWebViews[i].dwType).c_str());
			addBlock(
				m_lpWebViews[i].dwFlags,
				L"dwFlags = 0x%1!08X! = %2!ws!\r\n",
				m_lpWebViews[i].dwFlags.getData(),
				interpretprop::InterpretFlags(flagWebViewFlags, m_lpWebViews[i].dwFlags).c_str());
			addHeader(L"dwUnused = ");

			addBlockBytes(m_lpWebViews[i].dwUnused);

			addLine();
			addBlock(m_lpWebViews[i].cbData, L"cbData = 0x%1!08X!", m_lpWebViews[i].cbData.getData());

			addLine();
			switch (m_lpWebViews[i].dwType)
			{
			case WEBVIEWURL:
			{
				addHeader(L"wzURL = ");
				addBlock(m_lpWebViews[i].lpData, strings::BinToTextStringW(m_lpWebViews[i].lpData, false));
				break;
			}
			default:
				addHeader(L"lpData = ");
				addBlockBytes(m_lpWebViews[i].lpData);
				break;
			}
		}
	}
}