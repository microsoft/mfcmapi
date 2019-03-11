#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct WebViewPersist
	{
		blockT<DWORD> dwVersion;
		blockT<DWORD> dwType;
		blockT<DWORD> dwFlags;
		blockBytes dwUnused; // 7 DWORDs
		blockT<DWORD> cbData;
		blockBytes lpData;

		WebViewPersist(std::shared_ptr<binaryParser>& parser);
	};

	class WebViewPersistStream : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		DWORD m_cWebViews{};
		std::vector<std::shared_ptr<WebViewPersist>> m_lpWebViews;
	};
} // namespace smartview