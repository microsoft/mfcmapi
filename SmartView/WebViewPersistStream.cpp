#include "stdafx.h"
#include "..\stdafx.h"
#include "WebViewPersistStream.h"
#include "..\String.h"
#include "..\InterpretProp2.h"
#include "..\ParseProperty.h"
#include "..\ExtraPropTags.h"

WebViewPersistStream::WebViewPersistStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_cWebViews = 0;
	m_lpWebViews = NULL;
}

WebViewPersistStream::~WebViewPersistStream()
{
	if (m_lpWebViews && m_cWebViews)
	{
		ULONG i = 0;

		for (i = 0; i < m_cWebViews; i++)
		{
			delete[] m_lpWebViews[i].lpData;
		}
	}

	delete[] m_lpWebViews;
}

void WebViewPersistStream::Parse()
{
	// Run through the parser once to count the number of web view structs
	for (;;)
	{
		// Must have at least 2 bytes left to have another struct
		if (m_Parser.RemainingBytes() < sizeof(DWORD)* 11) break;
		m_Parser.Advance(sizeof(DWORD)* 10);
		DWORD cbData;
		m_Parser.GetDWORD(&cbData);

		// Must have at least cbData bytes left to be a valid flag
		if (m_Parser.RemainingBytes() < cbData) break;

		m_Parser.Advance(cbData);
		m_cWebViews++;
	}

	// Now we parse for real
	m_Parser.Rewind();

	DWORD cWebViews = m_cWebViews;
	if (cWebViews && cWebViews < _MaxEntriesSmall)
	{
		m_lpWebViews = new WebViewPersistStruct[cWebViews];
	}

	if (m_lpWebViews)
	{
		memset(m_lpWebViews, 0, sizeof(WebViewPersistStruct)*cWebViews);
		ULONG i = 0;

		for (i = 0; i < cWebViews; i++)
		{
			m_Parser.GetDWORD(&m_lpWebViews[i].dwVersion);
			m_Parser.GetDWORD(&m_lpWebViews[i].dwType);
			m_Parser.GetDWORD(&m_lpWebViews[i].dwFlags);
			m_Parser.GetBYTESNoAlloc(sizeof(m_lpWebViews[i].dwUnused), sizeof(m_lpWebViews[i].dwUnused), (LPBYTE)&m_lpWebViews[i].dwUnused);
			m_Parser.GetDWORD(&m_lpWebViews[i].cbData);
			m_Parser.GetBYTES(m_lpWebViews[i].cbData, _MaxBytes, &m_lpWebViews[i].lpData);
		}
	}
}

_Check_return_ wstring WebViewPersistStream::ToStringInternal()
{
	wstring szWebViewPersistStream = formatmessage(IDS_WEBVIEWSTREAMHEADER, m_cWebViews);
	if (m_lpWebViews && m_cWebViews)
	{
		ULONG i = 0;

		for (i = 0; i < m_cWebViews; i++)
		{
			wstring szVersion = InterpretFlags(flagWebViewVersion, m_lpWebViews[i].dwVersion);
			wstring szType = InterpretFlags(flagWebViewType, m_lpWebViews[i].dwType);
			wstring szFlags = InterpretFlags(flagWebViewFlags, m_lpWebViews[i].dwFlags);

			szWebViewPersistStream += formatmessage(
				IDS_WEBVIEWHEADER,
				i,
				m_lpWebViews[i].dwVersion, szVersion.c_str(),
				m_lpWebViews[i].dwType, szType.c_str(),
				m_lpWebViews[i].dwFlags, szFlags.c_str());

			SBinary sBin = { 0 };
			sBin.cb = sizeof(m_lpWebViews[i].dwUnused);
			sBin.lpb = (LPBYTE)&m_lpWebViews[i].dwUnused;
			szWebViewPersistStream += BinToHexString(&sBin, true);

			szWebViewPersistStream += formatmessage(IDS_WEBVIEWCBDATA, m_lpWebViews[i].cbData);

			switch (m_lpWebViews[i].dwType)
			{
			case WEBVIEWURL:
			{
				// Copy lpData to a new buffer and NULL terminate it in case it's not already.
				size_t cchData = m_lpWebViews[i].cbData / sizeof(WCHAR);
				WCHAR* lpwzTmp = new WCHAR[cchData + 1];
				if (lpwzTmp)
				{
					memcpy(lpwzTmp, m_lpWebViews[i].lpData, sizeof(WCHAR)* cchData);
					lpwzTmp[cchData] = NULL;
					szWebViewPersistStream += formatmessage(IDS_WEBVIEWURL);
					szWebViewPersistStream += lpwzTmp;
					delete[] lpwzTmp;
				}
				break;
			}
			default:
				sBin.cb = m_lpWebViews[i].cbData;
				sBin.lpb = m_lpWebViews[i].lpData;

				szWebViewPersistStream += formatmessage(IDS_WEBVIEWDATA);
				szWebViewPersistStream += BinToHexString(&sBin, true);
				break;
			}
		}
	}

	return szWebViewPersistStream;
}