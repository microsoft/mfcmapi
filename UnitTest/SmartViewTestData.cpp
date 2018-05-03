#include "stdafx.h"
#include "SmartViewTestData.h"
#include "SmartViewTestData_arp.h"
#include "SmartViewTestData_aei.h"

namespace SmartViewTestData
{
	vector<SmartViewTestData> g_smartViewTestData;

	void init()
	{
		SmartViewTestData_arp::init();
		SmartViewTestData_aei::init();
	}
}