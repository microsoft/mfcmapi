#include "stdafx.h"
#include "AppointmentRecurrencePattern.h"
#include "SmartView.h"
#include "String.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"
#include "Guids.h"

AppointmentRecurrencePattern::AppointmentRecurrencePattern(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	Init(cbBin, lpBin);
	m_RecurrencePattern = nullptr;
	m_ReaderVersion2 = 0;
	m_WriterVersion2 = 0;
	m_StartTimeOffset = 0;
	m_EndTimeOffset = 0;
	m_ExceptionCount = 0;
	m_ExceptionInfo = nullptr;
	m_ReservedBlock1Size = 0;
	m_ReservedBlock1 = nullptr;
	m_ExtendedException = nullptr;
	m_ReservedBlock2Size = 0;
	m_ReservedBlock2 = nullptr;
}

AppointmentRecurrencePattern::~AppointmentRecurrencePattern()
{
	delete m_RecurrencePattern;
	if (m_ExceptionCount && m_ExceptionInfo)
	{
		for (auto i = 0; i < m_ExceptionCount; i++)
		{
			delete[] m_ExceptionInfo[i].Subject;
			delete[] m_ExceptionInfo[i].Location;
		}
	}

	delete[] m_ExceptionInfo;
	delete[] m_ReservedBlock1;

	if (m_ExceptionCount && m_ExtendedException)
	{
		for (auto i = 0; i < m_ExceptionCount; i++)
		{
			delete[] m_ExtendedException[i].ChangeHighlight.Reserved;
			delete[] m_ExtendedException[i].ReservedBlockEE1;
			delete[] m_ExtendedException[i].WideCharSubject;
			delete[] m_ExtendedException[i].WideCharLocation;
			delete[] m_ExtendedException[i].ReservedBlockEE2;
		}
	}

	delete[] m_ExtendedException;
	delete[] m_ReservedBlock2;
}

void AppointmentRecurrencePattern::Parse()
{
	m_RecurrencePattern = new RecurrencePattern(static_cast<ULONG>(m_Parser.RemainingBytes()), m_Parser.GetCurrentAddress());

	if (m_RecurrencePattern)
	{
		m_RecurrencePattern->DisableJunkParsing();
		m_RecurrencePattern->EnsureParsed();
		m_Parser.Advance(m_RecurrencePattern->GetCurrentOffset());
	}

	m_Parser.GetDWORD(&m_ReaderVersion2);
	m_Parser.GetDWORD(&m_WriterVersion2);
	m_Parser.GetDWORD(&m_StartTimeOffset);
	m_Parser.GetDWORD(&m_EndTimeOffset);
	m_Parser.GetWORD(&m_ExceptionCount);

	if (m_ExceptionCount &&
		m_ExceptionCount == m_RecurrencePattern->m_ModifiedInstanceCount &&
		m_ExceptionCount < _MaxEntriesSmall)
	{
		m_ExceptionInfo = new ExceptionInfo[m_ExceptionCount];
		if (m_ExceptionInfo)
		{
			memset(m_ExceptionInfo, 0, sizeof(ExceptionInfo)* m_ExceptionCount);
			for (WORD i = 0; i < m_ExceptionCount; i++)
			{
				m_Parser.GetDWORD(&m_ExceptionInfo[i].StartDateTime);
				m_Parser.GetDWORD(&m_ExceptionInfo[i].EndDateTime);
				m_Parser.GetDWORD(&m_ExceptionInfo[i].OriginalStartDate);
				m_Parser.GetWORD(&m_ExceptionInfo[i].OverrideFlags);
				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
				{
					m_Parser.GetWORD(&m_ExceptionInfo[i].SubjectLength);
					m_Parser.GetWORD(&m_ExceptionInfo[i].SubjectLength2);
					if (m_ExceptionInfo[i].SubjectLength2 && m_ExceptionInfo[i].SubjectLength2 + 1 == m_ExceptionInfo[i].SubjectLength)
					{
						m_Parser.GetStringA(m_ExceptionInfo[i].SubjectLength2, &m_ExceptionInfo[i].Subject);
					}
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_MEETINGTYPE)
				{
					m_Parser.GetDWORD(&m_ExceptionInfo[i].MeetingType);
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_REMINDERDELTA)
				{
					m_Parser.GetDWORD(&m_ExceptionInfo[i].ReminderDelta);
				}
				if (m_ExceptionInfo[i].OverrideFlags & ARO_REMINDER)
				{
					m_Parser.GetDWORD(&m_ExceptionInfo[i].ReminderSet);
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					m_Parser.GetWORD(&m_ExceptionInfo[i].LocationLength);
					m_Parser.GetWORD(&m_ExceptionInfo[i].LocationLength2);
					if (m_ExceptionInfo[i].LocationLength2 && m_ExceptionInfo[i].LocationLength2 + 1 == m_ExceptionInfo[i].LocationLength)
					{
						m_Parser.GetStringA(m_ExceptionInfo[i].LocationLength2, &m_ExceptionInfo[i].Location);
					}
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_BUSYSTATUS)
				{
					m_Parser.GetDWORD(&m_ExceptionInfo[i].BusyStatus);
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_ATTACHMENT)
				{
					m_Parser.GetDWORD(&m_ExceptionInfo[i].Attachment);
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBTYPE)
				{
					m_Parser.GetDWORD(&m_ExceptionInfo[i].SubType);
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_APPTCOLOR)
				{
					m_Parser.GetDWORD(&m_ExceptionInfo[i].AppointmentColor);
				}
			}
		}
	}

	m_Parser.GetDWORD(&m_ReservedBlock1Size);
	m_Parser.GetBYTES(m_ReservedBlock1Size, _MaxBytes, &m_ReservedBlock1);

	if (m_ExceptionCount &&
		m_ExceptionCount == m_RecurrencePattern->m_ModifiedInstanceCount &&
		m_ExceptionCount < _MaxEntriesSmall &&
		m_ExceptionInfo)
	{
		m_ExtendedException = new ExtendedException[m_ExceptionCount];
		if (m_ExtendedException)
		{
			memset(m_ExtendedException, 0, sizeof(ExtendedException)* m_ExceptionCount);
			for (WORD i = 0; i < m_ExceptionCount; i++)
			{
				if (m_WriterVersion2 >= 0x0003009)
				{
					m_Parser.GetDWORD(&m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize);
					m_Parser.GetDWORD(&m_ExtendedException[i].ChangeHighlight.ChangeHighlightValue);
					if (m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
					{
						m_Parser.GetBYTES(m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize - sizeof(DWORD), _MaxBytes, &m_ExtendedException[i].ChangeHighlight.Reserved);
					}
				}

				m_Parser.GetDWORD(&m_ExtendedException[i].ReservedBlockEE1Size);
				m_Parser.GetBYTES(m_ExtendedException[i].ReservedBlockEE1Size, _MaxBytes, &m_ExtendedException[i].ReservedBlockEE1);

				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
					m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					m_Parser.GetDWORD(&m_ExtendedException[i].StartDateTime);
					m_Parser.GetDWORD(&m_ExtendedException[i].EndDateTime);
					m_Parser.GetDWORD(&m_ExtendedException[i].OriginalStartDate);
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
				{
					m_Parser.GetWORD(&m_ExtendedException[i].WideCharSubjectLength);
					if (m_ExtendedException[i].WideCharSubjectLength)
					{
						m_Parser.GetStringW(m_ExtendedException[i].WideCharSubjectLength, &m_ExtendedException[i].WideCharSubject);
					}
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					m_Parser.GetWORD(&m_ExtendedException[i].WideCharLocationLength);
					if (m_ExtendedException[i].WideCharLocationLength)
					{
						m_Parser.GetStringW(m_ExtendedException[i].WideCharLocationLength, &m_ExtendedException[i].WideCharLocation);
					}
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
					m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					m_Parser.GetDWORD(&m_ExtendedException[i].ReservedBlockEE2Size);
					m_Parser.GetBYTES(m_ExtendedException[i].ReservedBlockEE2Size, _MaxBytes, &m_ExtendedException[i].ReservedBlockEE2);

				}
			}
		}
	}

	m_Parser.GetDWORD(&m_ReservedBlock2Size);
	m_Parser.GetBYTES(m_ReservedBlock2Size, _MaxBytes, &m_ReservedBlock2);
}

_Check_return_ wstring AppointmentRecurrencePattern::ToStringInternal()
{
	wstring szARP;
	wstring szTmp;

	szARP = m_RecurrencePattern->ToString();

	szARP += formatmessage(IDS_ARPHEADER,
		m_ReaderVersion2,
		m_WriterVersion2,
		m_StartTimeOffset, RTimeToString(m_StartTimeOffset).c_str(),
		m_EndTimeOffset, RTimeToString(m_EndTimeOffset).c_str(),
		m_ExceptionCount);

	if (m_ExceptionCount && m_ExceptionInfo)
	{
		for (WORD i = 0; i < m_ExceptionCount; i++)
		{
			auto szOverrideFlags = InterpretFlags(flagOverrideFlags, m_ExceptionInfo[i].OverrideFlags);
			auto szExceptionInfo = formatmessage(IDS_ARPEXHEADER,
				i, m_ExceptionInfo[i].StartDateTime, RTimeToString(m_ExceptionInfo[i].StartDateTime).c_str(),
				m_ExceptionInfo[i].EndDateTime, RTimeToString(m_ExceptionInfo[i].EndDateTime).c_str(),
				m_ExceptionInfo[i].OriginalStartDate, RTimeToString(m_ExceptionInfo[i].OriginalStartDate).c_str(),
				m_ExceptionInfo[i].OverrideFlags, szOverrideFlags.c_str());

			if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
			{
				szExceptionInfo += formatmessage(IDS_ARPEXSUBJECT,
					i, m_ExceptionInfo[i].SubjectLength,
					m_ExceptionInfo[i].SubjectLength2,
					m_ExceptionInfo[i].Subject ? m_ExceptionInfo[i].Subject : "");
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_MEETINGTYPE)
			{
				auto szFlags = InterpretNumberAsStringNamedProp(m_ExceptionInfo[i].MeetingType, dispidApptStateFlags, const_cast<LPGUID>(&PSETID_Appointment));
				szExceptionInfo += formatmessage(IDS_ARPEXMEETINGTYPE,
					i, m_ExceptionInfo[i].MeetingType, szFlags.c_str());
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_REMINDERDELTA)
			{
				szExceptionInfo += formatmessage(IDS_ARPEXREMINDERDELTA,
					i, m_ExceptionInfo[i].ReminderDelta);
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_REMINDER)
			{
				szExceptionInfo += formatmessage(IDS_ARPEXREMINDERSET,
					i, m_ExceptionInfo[i].ReminderSet);
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
			{
				szExceptionInfo += formatmessage(IDS_ARPEXLOCATION,
					i, m_ExceptionInfo[i].LocationLength,
					m_ExceptionInfo[i].LocationLength2,
					m_ExceptionInfo[i].Location ? m_ExceptionInfo[i].Location : "");
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_BUSYSTATUS)
			{
				auto szFlags = InterpretNumberAsStringNamedProp(m_ExceptionInfo[i].BusyStatus, dispidBusyStatus, const_cast<LPGUID>(&PSETID_Appointment));
				szExceptionInfo += formatmessage(IDS_ARPEXBUSYSTATUS,
					i, m_ExceptionInfo[i].BusyStatus, szFlags.c_str());
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_ATTACHMENT)
			{
				szExceptionInfo += formatmessage(IDS_ARPEXATTACHMENT,
					i, m_ExceptionInfo[i].Attachment);
			}

			if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBTYPE)
			{
				szExceptionInfo += formatmessage(IDS_ARPEXSUBTYPE,
					i, m_ExceptionInfo[i].SubType);
			}
			if (m_ExceptionInfo[i].OverrideFlags & ARO_APPTCOLOR)
			{
				szExceptionInfo += formatmessage(IDS_ARPEXAPPOINTMENTCOLOR,
					i, m_ExceptionInfo[i].AppointmentColor);
			}

			szARP += szExceptionInfo;
		}
	}

	szARP += formatmessage(IDS_ARPRESERVED1,
		m_ReservedBlock1Size);
	if (m_ReservedBlock1Size)
	{
		SBinary sBin = { 0 };
		sBin.cb = m_ReservedBlock1Size;
		sBin.lpb = m_ReservedBlock1;
		szARP += BinToHexString(&sBin, true);
	}

	if (m_ExceptionCount && m_ExtendedException)
	{
		for (auto i = 0; i < m_ExceptionCount; i++)
		{
			wstring szExtendedException;
			if (m_WriterVersion2 >= 0x00003009)
			{
				auto szFlags = InterpretNumberAsStringNamedProp(m_ExtendedException[i].ChangeHighlight.ChangeHighlightValue, dispidChangeHighlight, const_cast<LPGUID>(&PSETID_Appointment));
				szExtendedException += formatmessage(IDS_ARPEXCHANGEHIGHLIGHT,
					i, m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize,
					m_ExtendedException[i].ChangeHighlight.ChangeHighlightValue, szFlags.c_str());

				if (m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
				{
					szExtendedException += formatmessage(IDS_ARPEXCHANGEHIGHLIGHTRESERVED,
						i);

					SBinary sBin = { 0 };
					sBin.cb = m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize - sizeof(DWORD);
					sBin.lpb = m_ExtendedException[i].ChangeHighlight.Reserved;
					szExtendedException += BinToHexString(&sBin, true);
					szExtendedException += L"\n"; // STRING_OK
				}
			}

			szExtendedException += formatmessage(IDS_ARPEXRESERVED1,
				i, m_ExtendedException[i].ReservedBlockEE1Size);
			if (m_ExtendedException[i].ReservedBlockEE1Size)
			{
				SBinary sBin = { 0 };
				sBin.cb = m_ExtendedException[i].ReservedBlockEE1Size;
				sBin.lpb = m_ExtendedException[i].ReservedBlockEE1;
				szExtendedException += BinToHexString(&sBin, true);
			}

			if (m_ExceptionInfo)
			{
				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT ||
					m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					szExtendedException += formatmessage(IDS_ARPEXDATETIME,
						i, m_ExtendedException[i].StartDateTime, RTimeToString(m_ExtendedException[i].StartDateTime).c_str(),
						m_ExtendedException[i].EndDateTime, RTimeToString(m_ExtendedException[i].EndDateTime).c_str(),
						m_ExtendedException[i].OriginalStartDate, RTimeToString(m_ExtendedException[i].OriginalStartDate).c_str());
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_SUBJECT)
				{
					szExtendedException += formatmessage(IDS_ARPEXWIDESUBJECT,
						i, m_ExtendedException[i].WideCharSubjectLength,
						m_ExtendedException[i].WideCharSubject ? m_ExtendedException[i].WideCharSubject : L"");
				}

				if (m_ExceptionInfo[i].OverrideFlags & ARO_LOCATION)
				{
					szExtendedException += formatmessage(IDS_ARPEXWIDELOCATION,
						i, m_ExtendedException[i].WideCharLocationLength,
						m_ExtendedException[i].WideCharLocation ? m_ExtendedException[i].WideCharLocation : L"");
				}
			}

			szExtendedException += formatmessage(IDS_ARPEXRESERVED1,
				i, m_ExtendedException[i].ReservedBlockEE2Size);
			if (m_ExtendedException[i].ReservedBlockEE2Size)
			{
				SBinary sBin = { 0 };
				sBin.cb = m_ExtendedException[i].ReservedBlockEE2Size;
				sBin.lpb = m_ExtendedException[i].ReservedBlockEE2;
				szExtendedException += BinToHexString(&sBin, true);
			}

			szARP += szExtendedException;
		}
	}

	szARP += formatmessage(IDS_ARPRESERVED2,
		m_ReservedBlock2Size);
	if (m_ReservedBlock2Size)
	{
		SBinary sBin = { 0 };
		sBin.cb = m_ReservedBlock2Size;
		sBin.lpb = m_ReservedBlock2;
		szARP += BinToHexString(&sBin, true);
	}

	return szARP;
}