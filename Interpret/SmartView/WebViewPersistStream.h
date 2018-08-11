#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	struct WebViewPersist
	{
		blockT<DWORD >dwVersion;
		blockT < DWORD> dwType;
		blockT < DWORD >dwFlags;
		blockBytes dwUnused; // 7 DWORDs
		blockT < DWORD >cbData;
		blockBytes lpData;
	};

	class WebViewPersistStream : public SmartViewParser
	{
	public:
		WebViewPersistStream();

	private:
		void Parse() override;
		void ParseBlocks() override;

		DWORD m_cWebViews;
		std::vector<WebViewPersist> m_lpWebViews;
	};
}