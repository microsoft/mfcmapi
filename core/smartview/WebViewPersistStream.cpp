#include <core/stdafx.h>
#include <core/smartview/WebViewPersistStream.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>

namespace smartview
{
	WebViewPersist::WebViewPersist(const std::shared_ptr<binaryParser>& parser)
	{
		dwVersion = blockT<DWORD>::parse(parser);
		dwType = blockT<DWORD>::parse(parser);
		dwFlags = blockT<DWORD>::parse(parser);
		dwUnused = blockBytes::parse(parser, 7 * sizeof(DWORD));
		cbData = blockT<DWORD>::parse(parser);
		lpData = blockBytes::parse(parser, *cbData, _MaxBytes);
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
				m_lpWebViews.emplace_back(std::make_shared<WebViewPersist>(parser));
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
			auto webView = create(L"Web View %1!d!", i);
			addChild(webView);
			webView->addChild(
				view->dwVersion,
				L"dwVersion = 0x%1!08X! = %2!ws!",
				view->dwVersion->getData(),
				flags::InterpretFlags(flagWebViewVersion, *view->dwVersion).c_str());
			webView->addChild(
				view->dwType,
				L"dwType = 0x%1!08X! = %2!ws!",
				view->dwType->getData(),
				flags::InterpretFlags(flagWebViewType, *view->dwType).c_str());
			webView->addChild(
				view->dwFlags,
				L"dwFlags = 0x%1!08X! = %2!ws!",
				view->dwFlags->getData(),
				flags::InterpretFlags(flagWebViewFlags, *view->dwFlags).c_str());
			webView->addLabeledChild(L"dwUnused =", view->dwUnused);

			webView->addChild(view->cbData, L"cbData = 0x%1!08X!", view->cbData->getData());

			switch (*view->dwType)
			{
			case WEBVIEWURL:
			{
				auto url = create(L"wzURL =");
				webView->addChild(url);
				url->addChild(view->lpData, view->lpData->toTextString(false));
				break;
			}
			default:
				webView->addLabeledChild(L"lpData =", view->lpData);
				break;
			}

			i++;
		}
	}
} // namespace smartview