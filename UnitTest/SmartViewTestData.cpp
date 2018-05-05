#include "stdafx.h"
#include "SmartViewTestData.h"
#include "SmartViewTestData_arp.h"
#include "SmartViewTestData_aei.h"
#include "resource.h"

namespace SmartViewTestData
{
	vector<SmartViewTestData> g_smartViewTestData;

	void init(HMODULE handle)
	{
		SmartViewTestData_aei::init(handle);
		SmartViewTestData_arp::init(handle);
	}

	wstring loadfile(HMODULE handle, int name)
	{
		auto rc = ::FindResource(handle, MAKEINTRESOURCE(name), MAKEINTRESOURCE(TEXTFILE));
		auto rcData = ::LoadResource(handle, rc);
		auto size = ::SizeofResource(handle, rc);
		auto data = static_cast<const char*>(::LockResource(rcData));
		auto ansi = std::string(data, size);
		return wstring(ansi.begin(), ansi.end());
	}
}