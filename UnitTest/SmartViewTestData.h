#pragma once

namespace SmartViewTest
{
	struct SmartViewTestData {
		wstring hex;
		__ParsingTypeEnum structType;
		wstring parsing;
	};

	extern vector<SmartViewTestData> const g_smartViewTestData;
}