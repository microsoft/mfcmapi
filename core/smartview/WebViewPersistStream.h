#pragma once
#include <core/smartview/block/smartViewParser.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct WebViewPersist
	{
		std::shared_ptr<blockT<DWORD>> dwVersion = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwType = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwFlags = emptyT<DWORD>();
		std::shared_ptr<blockBytes> dwUnused = emptyBB(); // 7 DWORDs
		std::shared_ptr<blockT<DWORD>> cbData = emptyT<DWORD>();
		std::shared_ptr<blockBytes> lpData = emptyBB();

		WebViewPersist(const std::shared_ptr<binaryParser>& parser);
	};

	class WebViewPersistStream : public smartViewParser
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::vector<std::shared_ptr<WebViewPersist>> m_lpWebViews;
	};
} // namespace smartview