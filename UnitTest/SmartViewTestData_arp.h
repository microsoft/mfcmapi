#pragma once
#include "SmartViewTestData.h"

namespace SmartViewTestData
{
	class SmartViewTestData_arp
	{
	public:
		static void init()
		{
			auto arp1 = SmartViewTestData
			{
				IDS_STAPPOINTMENTRECURRENCEPATTERN, false,
				L"043004300B2001000000C0210000020000000000000010000000212000001200000000000000040"
				L"00000806ADB0C40B9DB0C0008DC0CC056DC0C01000000806ADB0C80F4D80CA056DE0C0630000009"
				L"30000066030000A20300000100E66DDB0C226EDB0CE66DDB0C82020100000000000000000000000"
				L"4000000800100000000000000000000",
				L"Recurrence Pattern: \r\n"
				L"ReaderVersion: 0x3004\r\n"
				L"WriterVersion: 0x3004\r\n"
				L"RecurFrequency: 0x200B = IDC_RCEV_PAT_ORB_WEEKLY\r\n"
				L"PatternType: 0x0001 = rptWeek\r\n"
				L"CalendarType: 0x0000 = CAL_DEFAULT\r\n"
				L"FirstDateTime: 0x000021C0 = 8640\r\n"
				L"Period: 0x00000002 = 2\r\n"
				L"SlidingFlag: 0x00000000\r\n"
				L"PatternTypeSpecific.WeekRecurrencePattern: 0x00000010 = rdfThu\r\n"
				L"EndType: 0x00002021 = IDC_RCEV_PAT_ERB_END\r\n"
				L"OccurrenceCount: 0x00000012 = 18\r\n"
				L"FirstDOW: 0x00000000 = Sunday\r\n"
				L"DeletedInstanceCount: 0x00000004 = 4\r\n"
				L"DeletedInstanceDates[0]: 0x0CDB6A80 = 12:00:00.000 AM 2/17/2011\r\n"
				L"DeletedInstanceDates[1]: 0x0CDBB940 = 12:00:00.000 AM 3/3/2011\r\n"
				L"DeletedInstanceDates[2]: 0x0CDC0800 = 12:00:00.000 AM 3/17/2011\r\n"
				L"DeletedInstanceDates[3]: 0x0CDC56C0 = 12:00:00.000 AM 3/31/2011\r\n"
				L"ModifiedInstanceCount: 0x00000001 = 1\r\n"
				L"ModifiedInstanceDates[0]: 0x0CDB6A80 = 12:00:00.000 AM 2/17/2011\r\n"
				L"StartDate: 0x0CD8F480 = 215544960 = 12:00:00.000 AM 10/28/2010\r\n"
				L"EndDate: 0x0CDE56A0 = 215897760 = 12:00:00.000 AM 6/30/2011\r\n"
				L"Appointment Recurrence Pattern: \r\n"
				L"ReaderVersion2: 0x00003006\r\n"
				L"WriterVersion2: 0x00003009\r\n"
				L"StartTimeOffset: 0x00000366 = 870 = 02:30:00.000 PM 1/1/1601\r\n"
				L"EndTimeOffset: 0x000003A2 = 930 = 03:30:00.000 PM 1/1/1601\r\n"
				L"ExceptionCount: 0x0001\r\n"
				L"ExceptionInfo[0].StartDateTime: 0x0CDB6DE6 = 02:30:00.000 PM 2/17/2011\r\n"
				L"ExceptionInfo[0].EndDateTime: 0x0CDB6E22 = 03:30:00.000 PM 2/17/2011\r\n"
				L"ExceptionInfo[0].OriginalStartDate: 0x0CDB6DE6 = 02:30:00.000 PM 2/17/2011\r\n"
				L"ExceptionInfo[0].OverrideFlags: 0x0282 = ARO_MEETINGTYPE | ARO_SUBTYPE | ARO_EXCEPTIONAL_BODY\r\n"
				L"ExceptionInfo[0].MeetingType: 0x00000001 = asfMeeting\r\n"
				L"ExceptionInfo[0].SubType: 0x00000000\r\n"
				L"ReservedBlock1Size: 0x00000000\r\n"
				L"ExtendedException[0].ChangeHighlight.ChangeHighlightSize: 0x00000004\r\n"
				L"ExtendedException[0].ChangeHighlight.ChangeHighlightValue: 0x00000180 = BIT_CH_BODY | BIT_CH_CUSTOM\r\n"
				L"ExtendedException[0].ReservedBlockEE1Size: 0x00000000\r\n"
				L"ExtendedException[0].ReservedBlockEE2Size: 0x00000000\r\n"
				L"ReservedBlock2Size: 0x00000000" };

			g_smartViewTestData.push_back(arp1);
		}
	};
}