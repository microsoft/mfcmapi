#include <StdAfx.h>
#include <Interpret/SmartView/AppointmentRecurrencePattern.h>
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/Guids.h>

namespace smartview
{
	AppointmentRecurrencePattern::AppointmentRecurrencePattern() {}

	void AppointmentRecurrencePattern::Parse()
	{
		m_RecurrencePattern.Init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
		m_RecurrencePattern.DisableJunkParsing();
		m_RecurrencePattern.EnsureParsed();
		m_Parser.Advance(m_RecurrencePattern.GetCurrentOffset());

		m_ReaderVersion2 = m_Parser.GetBlock<DWORD>();
		m_WriterVersion2 = m_Parser.GetBlock<DWORD>();
		m_StartTimeOffset = m_Parser.GetBlock<DWORD>();
		m_EndTimeOffset = m_Parser.GetBlock<DWORD>();
		m_ExceptionCount = m_Parser.GetBlock<WORD>();

		if (m_ExceptionCount.getData() &&
			m_ExceptionCount.getData() == m_RecurrencePattern.m_ModifiedInstanceCount.getData() &&
			m_ExceptionCount.getData() < _MaxEntriesSmall)
		{
			for (WORD i = 0; i < m_ExceptionCount.getData(); i++)
			{
				ExceptionInfo exceptionInfo;
				exceptionInfo.StartDateTime = m_Parser.GetBlock<DWORD>();
				exceptionInfo.EndDateTime = m_Parser.GetBlock<DWORD>();
				exceptionInfo.OriginalStartDate = m_Parser.GetBlock<DWORD>();
				exceptionInfo.OverrideFlags = m_Parser.GetBlock<WORD>();
				if (exceptionInfo.OverrideFlags.getData() & ARO_SUBJECT)
				{
					exceptionInfo.SubjectLength = m_Parser.GetBlock<WORD>();
					exceptionInfo.SubjectLength2 = m_Parser.GetBlock<WORD>();
					if (exceptionInfo.SubjectLength2.getData() &&
						exceptionInfo.SubjectLength2.getData() + 1 == exceptionInfo.SubjectLength.getData())
					{
						exceptionInfo.Subject = m_Parser.GetBlockStringA(exceptionInfo.SubjectLength2.getData());
					}
				}

				if (exceptionInfo.OverrideFlags.getData() & ARO_MEETINGTYPE)
				{
					exceptionInfo.MeetingType = m_Parser.GetBlock<DWORD>();
				}

				if (exceptionInfo.OverrideFlags.getData() & ARO_REMINDERDELTA)
				{
					exceptionInfo.ReminderDelta = m_Parser.GetBlock<DWORD>();
				}
				if (exceptionInfo.OverrideFlags.getData() & ARO_REMINDER)
				{
					exceptionInfo.ReminderSet = m_Parser.GetBlock<DWORD>();
				}

				if (exceptionInfo.OverrideFlags.getData() & ARO_LOCATION)
				{
					exceptionInfo.LocationLength = m_Parser.GetBlock<WORD>();
					exceptionInfo.LocationLength2 = m_Parser.GetBlock<WORD>();
					if (exceptionInfo.LocationLength2.getData() &&
						exceptionInfo.LocationLength2.getData() + 1 == exceptionInfo.LocationLength.getData())
					{
						exceptionInfo.Location = m_Parser.GetBlockStringA(exceptionInfo.LocationLength2.getData());
					}
				}

				if (exceptionInfo.OverrideFlags.getData() & ARO_BUSYSTATUS)
				{
					exceptionInfo.BusyStatus = m_Parser.GetBlock<DWORD>();
				}

				if (exceptionInfo.OverrideFlags.getData() & ARO_ATTACHMENT)
				{
					exceptionInfo.Attachment = m_Parser.GetBlock<DWORD>();
				}

				if (exceptionInfo.OverrideFlags.getData() & ARO_SUBTYPE)
				{
					exceptionInfo.SubType = m_Parser.GetBlock<DWORD>();
				}

				if (exceptionInfo.OverrideFlags.getData() & ARO_APPTCOLOR)
				{
					exceptionInfo.AppointmentColor = m_Parser.GetBlock<DWORD>();
				}

				m_ExceptionInfo.push_back(exceptionInfo);
			}
		}

		m_ReservedBlock1Size = m_Parser.GetBlock<DWORD>();
		m_ReservedBlock1 = m_Parser.GetBlockBYTES(m_ReservedBlock1Size.getData(), _MaxBytes);

		if (m_ExceptionCount.getData() &&
			m_ExceptionCount.getData() == m_RecurrencePattern.m_ModifiedInstanceCount.getData() &&
			m_ExceptionCount.getData() < _MaxEntriesSmall && m_ExceptionInfo.size())
		{
			for (WORD i = 0; i < m_ExceptionCount.getData(); i++)
			{
				ExtendedException extendedException;

				std::vector<BYTE> ReservedBlockEE2;
				if (m_WriterVersion2.getData() >= 0x0003009)
				{
					extendedException.ChangeHighlight.ChangeHighlightSize = m_Parser.GetBlock<DWORD>();
					extendedException.ChangeHighlight.ChangeHighlightValue = m_Parser.GetBlock<DWORD>();
					if (extendedException.ChangeHighlight.ChangeHighlightSize.getData() > sizeof(DWORD))
					{
						extendedException.ChangeHighlight.Reserved = m_Parser.GetBlockBYTES(
							extendedException.ChangeHighlight.ChangeHighlightSize.getData() - sizeof(DWORD), _MaxBytes);
					}
				}

				extendedException.ReservedBlockEE1Size = m_Parser.GetBlock<DWORD>();
				extendedException.ReservedBlockEE1 =
					m_Parser.GetBlockBYTES(extendedException.ReservedBlockEE1Size.getData(), _MaxBytes);

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_SUBJECT ||
					m_ExceptionInfo[i].OverrideFlags.getData() & ARO_LOCATION)
				{
					extendedException.StartDateTime = m_Parser.GetBlock<DWORD>();
					extendedException.EndDateTime = m_Parser.GetBlock<DWORD>();
					extendedException.OriginalStartDate = m_Parser.GetBlock<DWORD>();
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_SUBJECT)
				{
					extendedException.WideCharSubjectLength = m_Parser.GetBlock<WORD>();
					if (extendedException.WideCharSubjectLength.getData())
					{
						extendedException.WideCharSubject =
							m_Parser.GetBlockStringW(extendedException.WideCharSubjectLength.getData());
					}
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_LOCATION)
				{
					extendedException.WideCharLocationLength = m_Parser.GetBlock<WORD>();
					if (extendedException.WideCharLocationLength.getData())
					{
						extendedException.WideCharLocation =
							m_Parser.GetBlockStringW(extendedException.WideCharLocationLength.getData());
					}
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_SUBJECT ||
					m_ExceptionInfo[i].OverrideFlags.getData() & ARO_LOCATION)
				{
					extendedException.ReservedBlockEE2Size = m_Parser.GetBlock<DWORD>();
					extendedException.ReservedBlockEE2 =
						m_Parser.GetBlockBYTES(extendedException.ReservedBlockEE2Size.getData(), _MaxBytes);
				}

				m_ExtendedException.push_back(extendedException);
			}
		}

		m_ReservedBlock2Size = m_Parser.GetBlock<DWORD>();
		m_ReservedBlock2 = m_Parser.GetBlockBYTES(m_ReservedBlock2Size.getData(), _MaxBytes);
	}

	void AppointmentRecurrencePattern::ParseBlocks()
	{
		addBlock(m_RecurrencePattern.getBlock());

		addLine();
		addHeader(L"Appointment Recurrence Pattern: \r\n");
		addBlock(m_ReaderVersion2, L"ReaderVersion2: 0x%1!08X!\r\n", m_ReaderVersion2.getData());
		addBlock(m_WriterVersion2, L"WriterVersion2: 0x%1!08X!\r\n", m_WriterVersion2.getData());
		addBlock(
			m_StartTimeOffset,
			L"StartTimeOffset: 0x%1!08X! = %1!d! = %2!ws!\r\n",
			m_StartTimeOffset.getData(),
			RTimeToString(m_StartTimeOffset.getData()).c_str());
		addBlock(
			m_EndTimeOffset,
			L"EndTimeOffset: 0x%1!08X! = %1!d! = %2!ws!\r\n",
			m_EndTimeOffset.getData(),
			RTimeToString(m_EndTimeOffset.getData()).c_str());
		addBlock(m_ExceptionCount, L"ExceptionCount: 0x%1!04X!\r\n", m_ExceptionCount.getData());

		if (m_ExceptionInfo.size())
		{
			for (WORD i = 0; i < m_ExceptionInfo.size(); i++)
			{
				addBlock(
					m_ExceptionInfo[i].StartDateTime,
					L"ExceptionInfo[%1!d!].StartDateTime: 0x%2!08X! = %3!ws!\r\n",
					i,
					m_ExceptionInfo[i].StartDateTime.getData(),
					RTimeToString(m_ExceptionInfo[i].StartDateTime.getData()).c_str());
				addBlock(
					m_ExceptionInfo[i].EndDateTime,
					L"ExceptionInfo[%1!d!].EndDateTime: 0x%2!08X! = %3!ws!\r\n",
					i,
					m_ExceptionInfo[i].EndDateTime.getData(),
					RTimeToString(m_ExceptionInfo[i].EndDateTime.getData()).c_str());
				addBlock(
					m_ExceptionInfo[i].OriginalStartDate,
					L"ExceptionInfo[%1!d!].OriginalStartDate: 0x%2!08X! = %3!ws!\r\n",
					i,
					m_ExceptionInfo[i].OriginalStartDate.getData(),
					RTimeToString(m_ExceptionInfo[i].OriginalStartDate.getData()).c_str());
				auto szOverrideFlags =
					interpretprop::InterpretFlags(flagOverrideFlags, m_ExceptionInfo[i].OverrideFlags.getData());
				addBlock(
					m_ExceptionInfo[i].OverrideFlags,
					L"ExceptionInfo[%1!d!].OverrideFlags: 0x%2!04X! = %3!ws!\r\n",
					i,
					m_ExceptionInfo[i].OverrideFlags.getData(),
					szOverrideFlags.c_str());

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_SUBJECT)
				{
					addBlock(
						m_ExceptionInfo[i].SubjectLength,
						L"ExceptionInfo[%1!d!].SubjectLength: 0x%2!04X! = %2!d!\r\n",
						i,
						m_ExceptionInfo[i].SubjectLength.getData());
					addBlock(
						m_ExceptionInfo[i].SubjectLength2,
						L"ExceptionInfo[%1!d!].SubjectLength2: 0x%2!04X! = %2!d!\r\n",
						i,
						m_ExceptionInfo[i].SubjectLength2.getData());

					addBlock(
						m_ExceptionInfo[i].Subject,
						L"ExceptionInfo[%1!d!].Subject: \"%2!hs!\"\r\n",
						i,
						m_ExceptionInfo[i].Subject.getData().c_str());
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_MEETINGTYPE)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						m_ExceptionInfo[i].MeetingType.getData(),
						dispidApptStateFlags,
						const_cast<LPGUID>(&guid::PSETID_Appointment));
					addBlock(
						m_ExceptionInfo[i].MeetingType,
						L"ExceptionInfo[%1!d!].MeetingType: 0x%2!08X! = %3!ws!\r\n",
						i,
						m_ExceptionInfo[i].MeetingType.getData(),
						szFlags.c_str());
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_REMINDERDELTA)
				{
					addBlock(
						m_ExceptionInfo[i].ReminderDelta,
						L"ExceptionInfo[%1!d!].ReminderDelta: 0x%2!08X!\r\n",
						i,
						m_ExceptionInfo[i].ReminderDelta.getData());
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_REMINDER)
				{
					addBlock(
						m_ExceptionInfo[i].ReminderSet,
						L"ExceptionInfo[%1!d!].ReminderSet: 0x%2!08X!\r\n",
						i,
						m_ExceptionInfo[i].ReminderSet.getData());
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_LOCATION)
				{
					addBlock(
						m_ExceptionInfo[i].LocationLength,
						L"ExceptionInfo[%1!d!].LocationLength: 0x%2!04X! = %2!d!\r\n",
						i,
						m_ExceptionInfo[i].LocationLength.getData());
					addBlock(
						m_ExceptionInfo[i].LocationLength2,
						L"ExceptionInfo[%1!d!].LocationLength2: 0x%2!04X! = %2!d!\r\n",
						i,
						m_ExceptionInfo[i].LocationLength2.getData());
					addBlock(
						m_ExceptionInfo[i].Location,
						L"ExceptionInfo[%1!d!].Location: \"%2!hs!\"\r\n",
						i,
						m_ExceptionInfo[i].Location.getData().c_str());
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_BUSYSTATUS)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						m_ExceptionInfo[i].BusyStatus.getData(),
						dispidBusyStatus,
						const_cast<LPGUID>(&guid::PSETID_Appointment));
					addBlock(
						m_ExceptionInfo[i].BusyStatus,
						L"ExceptionInfo[%1!d!].BusyStatus: 0x%2!08X! = %3!ws!\r\n",
						i,
						m_ExceptionInfo[i].BusyStatus.getData(),
						szFlags.c_str());
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_ATTACHMENT)
				{
					addBlock(
						m_ExceptionInfo[i].Attachment,
						L"ExceptionInfo[%1!d!].Attachment: 0x%2!08X!\r\n",
						i,
						m_ExceptionInfo[i].Attachment.getData());
				}

				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_SUBTYPE)
				{
					addBlock(
						m_ExceptionInfo[i].SubType,
						L"ExceptionInfo[%1!d!].SubType: 0x%2!08X!\r\n",
						i,
						m_ExceptionInfo[i].SubType.getData());
				}
				if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_APPTCOLOR)
				{
					addBlock(
						m_ExceptionInfo[i].AppointmentColor,
						L"ExceptionInfo[%1!d!].AppointmentColor: 0x%2!08X!\r\n",
						i,
						m_ExceptionInfo[i].AppointmentColor.getData());
				}
			}
		}

		addBlock(m_ReservedBlock1Size, L"ReservedBlock1Size: 0x%1!08X!\r\n", m_ReservedBlock1Size.getData());
		if (m_ReservedBlock1Size.getData())
		{
			addBlockBytes(m_ReservedBlock1);
		}

		if (m_ExtendedException.size())
		{
			for (size_t i = 0; i < m_ExtendedException.size(); i++)
			{
				if (m_WriterVersion2.getData() >= 0x00003009)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						m_ExtendedException[i].ChangeHighlight.ChangeHighlightValue.getData(),
						dispidChangeHighlight,
						const_cast<LPGUID>(&guid::PSETID_Appointment));
					addBlock(
						m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize,
						L"ExtendedException[%1!d!].ChangeHighlight.ChangeHighlightSize: 0x%2!08X!\r\n",
						i,
						m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize.getData());
					addBlock(
						m_ExtendedException[i].ChangeHighlight.ChangeHighlightValue,
						L"ExtendedException[%1!d!].ChangeHighlight.ChangeHighlightValue: 0x%2!08X! = %3!ws!\r\n",
						i,
						m_ExtendedException[i].ChangeHighlight.ChangeHighlightValue.getData(),
						szFlags.c_str());

					if (m_ExtendedException[i].ChangeHighlight.ChangeHighlightSize.getData() > sizeof(DWORD))
					{
						addHeader(L"ExtendedException[%1!d!].ChangeHighlight.Reserved:", i);

						addBlockBytes(m_ExtendedException[i].ChangeHighlight.Reserved);
					}
				}

				addBlock(
					m_ExtendedException[i].ReservedBlockEE1Size,
					L"ExtendedException[%1!d!].ReservedBlockEE1Size: 0x%2!08X!\r\n",
					i,
					m_ExtendedException[i].ReservedBlockEE1Size.getData());
				if (m_ExtendedException[i].ReservedBlockEE1.getData().size())
				{
					addBlockBytes(m_ExtendedException[i].ReservedBlockEE1);
				}

				if (i < m_ExceptionInfo.size())
				{
					if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_SUBJECT ||
						m_ExceptionInfo[i].OverrideFlags.getData() & ARO_LOCATION)
					{
						addBlock(
							m_ExtendedException[i].StartDateTime,
							L"ExtendedException[%1!d!].StartDateTime: 0x%2!08X! = %3\r\n",
							i,
							m_ExtendedException[i].StartDateTime.getData(),
							RTimeToString(m_ExtendedException[i].StartDateTime.getData()).c_str());
						addBlock(
							m_ExtendedException[i].EndDateTime,
							L"ExtendedException[%1!d!].EndDateTime: 0x%2!08X! = %3!ws!\r\n",
							i,
							m_ExtendedException[i].EndDateTime.getData(),
							RTimeToString(m_ExtendedException[i].EndDateTime.getData()).c_str());
						addBlock(
							m_ExtendedException[i].OriginalStartDate,
							L"ExtendedException[%1!d!].OriginalStartDate: 0x%2!08X! = %3!ws!\r\n",
							i,
							m_ExtendedException[i].OriginalStartDate.getData(),
							RTimeToString(m_ExtendedException[i].OriginalStartDate.getData()).c_str());
					}

					if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_SUBJECT)
					{
						addBlock(
							m_ExtendedException[i].WideCharSubjectLength,
							L"ExtendedException[%1!d!].WideCharSubjectLength: 0x%2!08X! = %2!d!\r\n",
							i,
							m_ExtendedException[i].WideCharSubjectLength.getData());
						addBlock(
							m_ExtendedException[i].WideCharSubject,
							L"ExtendedException[%1!d!].WideCharSubject: \"%2!ws!\"\r\n",
							i,
							m_ExtendedException[i].WideCharSubject.getData().c_str());
					}

					if (m_ExceptionInfo[i].OverrideFlags.getData() & ARO_LOCATION)
					{
						addBlock(
							m_ExtendedException[i].WideCharLocationLength,
							L"ExtendedException[%1!d!].WideCharLocationLength: 0x%2!08X! = %2!d!\r\n",
							i,
							m_ExtendedException[i].WideCharLocationLength.getData());
						addBlock(
							m_ExtendedException[i].WideCharLocation,
							L"ExtendedException[%1!d!].WideCharLocation: \"%2!ws!\"\r\n",
							i,
							m_ExtendedException[i].WideCharLocation.getData().c_str());
					}
				}

				addBlock(
					m_ExtendedException[i].ReservedBlockEE2Size,
					L"ExtendedException[%1!d!].ReservedBlockEE2Size: 0x%2!08X!\r\n",
					i,
					m_ExtendedException[i].ReservedBlockEE2Size.getData());
				if (m_ExtendedException[i].ReservedBlockEE2Size.getData())
				{
					addBlockBytes(m_ExtendedException[i].ReservedBlockEE2);
				}
			}
		}

		addBlock(m_ReservedBlock2Size, L"ReservedBlock2Size: 0x%1!08X!", m_ReservedBlock2Size.getData());
		if (m_ReservedBlock2Size.getData())
		{
			addBlockBytes(m_ReservedBlock2);
		}
	}
}