#pragma once

namespace SmartViewTestData
{
	struct SmartViewTestResource {
		__ParsingTypeEnum structType;
		bool parseAll;
		DWORD hex;
		DWORD expected;
	};

	struct SmartViewTestData {
		__ParsingTypeEnum structType;
		bool parseAll;
		wstring testName;
		wstring hex;
		wstring expected;
	};

	vector<SmartViewTestData> loadTestData(HMODULE handle, std::initializer_list<SmartViewTestResource> resources);
}