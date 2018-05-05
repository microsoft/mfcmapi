#pragma once

namespace SmartViewTestData
{
	struct SmartViewTestData {
		__ParsingTypeEnum structType;
		bool parseAll;
		wstring hex;
		wstring expected;
	};

	vector<SmartViewTestData> getTestData(HMODULE handle);
}