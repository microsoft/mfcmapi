#include <core/stdafx.h>
#include <core/smartview/WebViewPersistStream.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>

namespace smartview
{
	void WebViewPersist::parse()
	{
		dwVersion = blockT<DWORD>::parse(parser);
		dwType = blockT<DWORD>::parse(parser);
		dwFlags = blockT<DWORD>::parse(parser);
		dwUnused = blockBytes::parse(parser, 7 * sizeof(DWORD));
		cbData = blockT<DWORD>::parse(parser);
		lpData = blockBytes::parse(parser, *cbData, _MaxBytes);
	}

	void WebViewPersist::parseBlocks()
	{
		addChild(
			dwVersion,
			L"dwVersion = 0x%1!08X! = %2!ws!",
			dwVersion->getData(),
			flags::InterpretFlags(flagWebViewVersion, *dwVersion).c_str());
		addChild(
			dwType,
			L"dwType = 0x%1!08X! = %2!ws!",
			dwType->getData(),
			flags::InterpretFlags(flagWebViewType, *dwType).c_str());
		addChild(
			dwFlags,
			L"dwFlags = 0x%1!08X! = %2!ws!",
			dwFlags->getData(),
			flags::InterpretFlags(flagWebViewFlags, *dwFlags).c_str());
		addLabeledChild(L"dwUnused =", dwUnused);

		addChild(cbData, L"cbData = 0x%1!08X!", cbData->getData());

		switch (*dwType)
		{
		case WEBVIEWURL:
		{
			auto url = create(L"wzURL =");
			addChild(url);
			url->addChild(lpData, lpData->toTextString(false));
			break;
		}
		default:
			addLabeledChild(L"lpData =", lpData);
			break;
		}
	}

	void WebViewPersistStream::parse()
	{
		auto cWebViews = 0;

		// Run through the parser once to count the number of web view structs
		for (;;)
		{
			// Must have at least 2 bytes left to have another struct
			if (parser->getSize() < sizeof(DWORD) * 11) break;
			parser->advance(sizeof(DWORD) * 10);
			const auto cbData = blockT<DWORD>::parse(parser);

			// Must have at least cbData bytes left to be a valid flag
			if (parser->getSize() < *cbData) break;

			parser->advance(*cbData);
			cWebViews++;
		}

		// Now we parse for real
		parser->rewind();

		if (cWebViews && cWebViews < _MaxEntriesSmall)
		{
			m_lpWebViews.reserve(cWebViews);
			for (auto i = 0; i < cWebViews; i++)
			{
				m_lpWebViews.emplace_back(block::parse<WebViewPersist>(parser, false));
			}
		}
	}

	void WebViewPersistStream::parseBlocks()
	{
		setText(L"Web View Persistence Object Stream");
		addSubHeader(L"cWebViews = %1!d!", m_lpWebViews.size());
		auto i = 0;
		for (const auto& view : m_lpWebViews)
		{
			addChild(view, L"Web View %1!d!", i++);
		}
	}
} // namespace smartview