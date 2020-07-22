#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class WebViewPersist : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> dwVersion = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwType = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwFlags = emptyT<DWORD>();
		std::shared_ptr<blockBytes> dwUnused = emptyBB(); // 7 DWORDs
		std::shared_ptr<blockT<DWORD>> cbData = emptyT<DWORD>();
		std::shared_ptr<blockBytes> lpData = emptyBB();
	};

	class WebViewPersistStream : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::vector<std::shared_ptr<WebViewPersist>> m_lpWebViews;
	};
} // namespace smartview