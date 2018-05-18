#include "stdafx.h"
#include "AppointmentRecurrencePattern.h"
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/Guids.h>

AppointmentRecurrencePattern::AppointmentRecurrencePattern()
{
	m_ReaderVersion2 = 0;
	m_WriterVersion2 = 0;
	m_StartTimeOffset = 0;
	m_EndTimeOffset = 0;
	m_ExceptionCount = 0;
	m_ReservedBlock1Size = 0;
	m_ReservedBlock2Size = 0;
}

void AppointmentRecurrencePattern::Parse()
{
	m_RecurrencePattern.Init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
	m_RecurrencePattern.DisableJunkParsing();
	m_RecurrencePattern.EnsureParsed();
	m_Parser.Advance(m_RecurrencePattern.GetCurrentOffset());

	m_ReaderVersion2 = m_Parser.Get<DWORD>();
	m_WriterVersion2 = m_Parser.Get<DWORD>();
	m_StartTimeOffset = m_Parser.Get<DWORD>();
	m_EndTimeOffset = m_Parser.Get<DWORD>();
	m_ExceptionCount = m_Parser.Get<WORD>();

	if (m_ExceptionCount &&
		m_ExceptionCount == m_RecurrencePattern.m_ModifiedInstanceCount &&
		m_ExceptionCount < _MaxEntriesSmall)
	{
		for (WORD i = 0; i < m_ExceptionCount; i++)
		{
			ExceptionInfo exceptionInfo;
			exceptionInfo.StartDateTime = m_Parser.Get<DWORD>();
			exceptionInfo.EndDateTime = m_Parser.Get<DWORD>();
			exceptionInfo.OriginalStartDate = m_Parser.Get<DWORD>();
			exceptionInfo.OverrideFlags = m_Parser.Get<WORD>();
			if (exceptionInfo.OverrideFlags & ARO_SUBJECT)
			{
				exceptionInfo.SubjectLength = m_Parser.Get<WORD>();
				exceptionInfo.SubjectLength2 = m_Parser.Get<WORD>();
				if (exceptionInfo.SubjectLength2 && exceptionInfo.SubjectLength2 + 1 == exceptionInfo.SubjectLength)
				{
					exceptionInfo.Subject = m_Parser.GetStringA(exceptionInfo.SubjectLength2);
				}
			}

			if (exceptionInfo.OverrideFlags & ARO_MEETINGTYPE)
			{
				exceptionInfo.MeetingType = m_Parser.Get<DWORD>();
			}

			if (exceptionInfo.OverrideFlags & ARO_REMINDERDELTA)
			{
				exceptionInfo.ReminderDelta = m_Parser.Get<DWORD>();
			}
			if (exceptionInfo.OverrideFlags & ARO_REMINDER)
			{
				exceptionInfo.ReminderSet = m_Parser.Get<DWORD>();
			}

			if (exceptionInfo.OverrideFlags & ARO_LOCATION)
			{
				exceptionInfo.LocationLength = m_Parser.Get<WORD>();
				exceptionInfo.LocationLength2 = m_Parser.Get<WORD>();
				if (exceptionInfo.LocationLength2 && exceptionInfo.LocationLength2 + 1 == exceptionInfo.LocationLength)
				{
					exceptionInfo.Location = m_Parser.GetStringA(exceptionInfo.LocationLength2);
				}
			}

			if (exceptionInfo.OverrideFlags & ARO_BUSYSTATUS)
			{
				exceptionInfo.BusyStatus = m_Parser.Get<DWORD>();
			}

			if (exceptionInfo.OverrideFlags & ARO_ATTACHMENT)
			{
				exceptionInfo.Attachment = m_Parser.Get<DWORD>();
			}

			if (exceptionInfo.OverrideFlags & ARO_SUBTYPE)
			{
				exceptionInfo.SubType = m_Parser.Get<DWORD>();
			}

			if (exceptionInfo.OverrideFlags & ARO_APPTCOLOR)
			{
				exceptionInfo.AppointmentColor = m_Parser.Get<DWORD>();
			}

			m_ExceptionInfo.push_back(exceptionInfo);
		}
	}

	m_ReservedBlock1Size = m_Parser.Get<DWORD>();
	m_ReservedBlock1 = m_Parser.GetBYTES(m_ReservedBlock1Size, _MaxBytes);

	if (m_ExceptionCount &&
		m_ExceptionCount == m_RecurrencePattern.m_ModifiedInstanceCount &&
		m_ExceptionCount < _MaxEntriesSmall &&
		m_ExceptionInfo.size())
	{
		for (WORD i = 0; i < m_ExceptionCount; i++)
		{
			ExtendedException extendedException;
			extendedException.ReservedBlockEE2Size = 0;
			extendedException.ReservedBlockEE1Size = 0;
			extendedException.StartDateTime = 0;
			extendedException.EndDateTime = 0;
			extendedException.OriginalStartDate = 0;
			extendedException.WideCharSubjectLength = 0;
			extendedException.WideCharLocationLength = 0;
			extendedException.ReservedBlockEE2Size = 0;

			vector<BYTE> ReservedBlockEE2;
			if (m_WriterVersion2 >= 0x0003009)
			{
				extendedException.ChangeHighlight.ChangeHighlightSize = m_Parser.Get<DWORD>();
				extendedException.ChangeHighlight.ChangeHighlightValue = m_Parser.Get<DWORD>();
				if (extendedException.ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
				{
					extendedException.ChangeHighlight.Reserved = m_Parser.GetBYTES(extendedException.ChangeHighlight.ChangeHighlightSize - sizeof(DWORD), _MaxBytes);
				}
			}

			extendedException.ReservedBlockEE1Size = m_Parser.Get<DWORD>();
			extendedException.ReservedBlockEE1 = m_Parser.GetBYTES(extendedException.ReservedBlockEE1Size, _MaxBytes);

			if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
				m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				extendedException.StartDateTime = m_Parser.Get<DWORD>();
				extendedException.EndDateTime = m_Parser.Get<DWORD>();
				extendedException.OriginalStartDate = m_Parser.Get<DWORD>();
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
			{
				extendedException.WideCharSubjectLength = m_Parser.Get<WORD>();
				if (extendedException.WideCharSubjectLength)
				{
					extendedException.WideCharSubject = m_Parser.GetStringW(extendedException.WideCharSubjectLength);
				}
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				extendedException.WideCharLocationLength = m_Parser.Get<WORD>();
				if (extendedException.WideCharLocationLength)
				{
					extendedException.WideCharLocation = m_Parser.GetStringW(extendedException.WideCharLocationLength);
				}
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
				m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				extendedException.ReservedBlockEE2Size = m_Parser.Get<DWORD>();
				extendedException.ReservedBlockEE2 = m_Parser.GetBYTES(extendedException.ReservedBlockEE2Size, _MaxBytes);
			}

			m_ExtendedException.push_back(extendedException);
		}
	}

	m_ReservedBlock2Size = m_Parser.Get<DWORD>();
	m_ReservedBlock2 = m_Parser.GetBYTES(m_ReservedBlock2Size, _MaxBytes);
}

_Check_return_ wstring AppointmentRecurrencePattern::ToStringInternal()
{
	wstring szARP;
	wstring szTmp;

	szARP = m_RecurrencePattern.ToString();

	szARP += strings::formatmessage(IDS_ARPHEADER,
		m_ReaderVersion2,
		m_WriterVersion2,
		m_StartTimeOffset, RTimeToString(m_StartTimeOffset).c_str(),
		m_EndTimeOffset, RTimeToString(m_EndTimeOffset).c_str(),
		m_ExceptionCount);

	if (m_ExceptionInfo.size())
	{
		for (WORD i = 0; i < m_ExceptionInfo.size(); i++)
		{
			auto szOverrideFlags = InterpretFlags(flagOverrideFlags, m_ExceptionInfo[i].OverrideFlags);
			auto szExceptionInfo = strings::formatmessage(IDS_ARPEXHEADER,
				i, m_ExceptionInfo[i].StartDateTime, RTimeToString(m_ExceptionInfo[i].StartDateTime).c_str(),
				m_ExceptionInfo[i].EndDateTime, RTimeToString(m_ExceptionInfo[i].EndDateTime).c_str(),
				m_ExceptionInfo[i].OriginalStartDate, RTimeToString(m_ExceptionInfo[i].OriginalStartDate).c_str(),
				m_ExceptionInfo[i].OverrideFlags, szOverrideFlags.c_str());

			if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
			{
				szExceptionInfo += strings::formatmessage(IDS_ARPEXSUBJECT,
					i, m_ExceptionInfo[i].SubjectLength,
					m_ExceptionInfo[i].SubjectLength2,
					m_ExceptionInfo[i].Subject.c_str());
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_MEETINGTYPE)
			{
				auto szFlags = InterpretNumberAsStringNamedProp(m_ExceptionInfo[i].MeetingType, dispidApptStateFlags, const_cast<LPGUID>(&PSETID_Appointment));
				szExceptionInfo += strings::formatmessage(IDS_ARPEXMEETINGTYPE,
					i, m_ExceptionInfo[i].MeetingType, szFlags.c_str());
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_REMINDERDELTA)
			{
				szExceptionInfo += strings::formatmessage(IDS_ARPEXREMINDERDELTA,
					i, m_ExceptionInfo[i].ReminderDelta);
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_REMINDER)
			{
				szExceptionInfo += strings::formatmessage(IDS_ARPEXREMINDERSET,
					i, m_ExceptionInfo[i].ReminderSet);
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				szExceptionInfo += strings::formatmessage(IDS_ARPEXLOCATION,
					i, m_ExceptionInfo[i].LocationLength,
					m_ExceptionInfo[i].LocationLength2,
					m_ExceptionInfo[i].Location.c_str());
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_BUSYSTATUS)
			{
				auto szFlags = InterpretNumberAsStringNamedProp(m_ExceptionInfo[i].BusyStatus, dispidBusyStatus, const_cast<LPGUID>(&PSETID_Appointment));
				szExceptionInfo += strings::formatmessage(IDS_ARPEXBUSYSTATUS,
					i, m_ExceptionInfo[i].BusyStatus, szFlags.c_str());
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_ATTACHMENT)
			{
				szExceptionInfo += strings::formatmessage(IDS_ARPEXATTACHMENT,
					i, m_ExceptionInfo[i].Attachment);
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBTYPE)
			{
				szExceptionInfo += strings::formatmessage(IDS_ARPEXSUBTYPE,
					i, m_ExceptionInfo[i].SubType);
			}
			if (m_ExceptionInfo[i].OverrideFlags & ARO_APPTCOLOR)
			{
				szExceptionInfo += strings::formatmessage(IDS_ARPEXAPPOINTMENTCOLOR,
					i, m_ExceptionInfo[i].AppointmentColor);
			}

			szARP += szExceptionInfo;
		}
	}

	szARP += strings::formatmessage(IDS_ARPRESERVED1,
		m_ReservedBlock1Size);
	if (m_ReservedBlock1Size)
	{
		szARP += strings::BinToHexString(m_ReservedBlock1, true);
	}

	if (m_ExtendedException.size())
	{
		for (size_t i = 0; i < m_ExtendedException.size(); i++)
		{
			wstring szExtendedException;
			if (m_WriterVersion2 >= 0x00003009)
			{
				auto szFlags = InterpretNumberAsStringNamedProp(m_ExtendedException[i].ChangeHighlight.ChangeHighlightValue, dispidChangeHighlight, const_cast<LPGUID>(&PSETID_Appointment));
				szExtendedException += strings::formatmessage(IDS_ARPEXCHANGEHIGHLIGHT,
					i, m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize,
					m_ExtendedException[i].ChangeHighlight.ChangeHighlightValue, szFlags.c_str());

				if (m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
				{
					szExtendedException += strings::formatmessage(IDS_ARPEXCHANGEHIGHLIGHTRESERVED,
						i);

					szExtendedException += strings::BinToHexString(m_ExtendedException[i].ChangeHighlight.Reserved, true);
					szExtendedException += L"\n"; // STRING_OK
				}
			}

			szExtendedException += strings::formatmessage(IDS_ARPEXRESERVED1,
				i, m_ExtendedException[i].ReservedBlockEE1Size);
			if (m_ExtendedException[i].ReservedBlockEE1.size())
			{
				szExtendedException += strings::BinToHexString(m_ExtendedException[i].ReservedBlockEE1, true);
			}

			if (i < m_ExceptionInfo.size())
			{
				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
					m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					szExtendedException += strings::formatmessage(IDS_ARPEXDATETIME,
						i, m_ExtendedException[i].StartDateTime, RTimeToString(m_ExtendedException[i].StartDateTime).c_str(),
						m_ExtendedException[i].EndDateTime, RTimeToString(m_ExtendedException[i].EndDateTime).c_str(),
						m_ExtendedException[i].OriginalStartDate, RTimeToString(m_ExtendedException[i].OriginalStartDate).c_str());
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
				{
					szExtendedException += strings::formatmessage(IDS_ARPEXWIDESUBJECT,
						i, m_ExtendedException[i].WideCharSubjectLength,
						m_ExtendedException[i].WideCharSubject.c_str());
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					szExtendedException += strings::formatmessage(IDS_ARPEXWIDELOCATION,
						i, m_ExtendedException[i].WideCharLocationLength,
						m_ExtendedException[i].WideCharLocation.c_str());
				}
			}

			szExtendedException += strings::formatmessage(IDS_ARPEXRESERVED2,
				i, m_ExtendedException[i].ReservedBlockEE2Size);
			if (m_ExtendedException[i].ReservedBlockEE2Size)
			{
				szExtendedException += strings::BinToHexString(m_ExtendedException[i].ReservedBlockEE2, true);
			}

			szARP += szExtendedException;
		}
	}

	szARP += strings::formatmessage(IDS_ARPRESERVED2,
		m_ReservedBlock2Size);
	if (m_ReservedBlock2Size)
	{
		szARP += strings::BinToHexString(m_ReservedBlock2, true);
	}

	return szARP;
}