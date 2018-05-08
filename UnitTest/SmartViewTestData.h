#pragma once

namespace SmartViewTestData
{
	struct SmartViewTestData {
		__ParsingTypeEnum structType;
		bool parseAll;
		wstring testName;
		wstring hex;
		wstring expected;
	};

	vector<SmartViewTestData> getTestData(HMODULE handle);
}