#pragma once

namespace SmartViewTestData
{
	struct SmartViewTestData {
		__ParsingTypeEnum structType;
		bool parseAll;
		wstring hex;
		wstring expected;
	};

	extern vector<SmartViewTestData> g_smartViewTestData;
	void init(HMODULE handle);
}