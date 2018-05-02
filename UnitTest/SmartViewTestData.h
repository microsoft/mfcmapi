#pragma once

namespace SmartViewTest
{
	struct SmartViewTestData {
		__ParsingTypeEnum structType;
		bool parseAll;
		wstring hex;
		wstring parsing;
	};

	extern vector<SmartViewTestData> const g_smartViewTestData;
}