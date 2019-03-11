#include <core/stdafx.h>
#include <core/smartview/AppointmentRecurrencePattern.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/guid.h>

namespace smartview
{
	ExceptionInfo::ExceptionInfo(std::shared_ptr<binaryParser>& parser)
	{
		StartDateTime.parse<DWORD>(parser);
		EndDateTime.parse<DWORD>(parser);
		OriginalStartDate.parse<DWORD>(parser);
		OverrideFlags.parse<WORD>(parser);
		if (OverrideFlags & ARO_SUBJECT)
		{
			SubjectLength.parse<WORD>(parser);
			SubjectLength2.parse<WORD>(parser);
			if (SubjectLength2 && SubjectLength2 + 1 == SubjectLength)
			{
				Subject.parse(parser, SubjectLength2);
			}
		}

		if (OverrideFlags & ARO_MEETINGTYPE)
		{
			MeetingType.parse<DWORD>(parser);
		}

		if (OverrideFlags & ARO_REMINDERDELTA)
		{
			ReminderDelta.parse<DWORD>(parser);
		}
		if (OverrideFlags & ARO_REMINDER)
		{
			ReminderSet.parse<DWORD>(parser);
		}

		if (OverrideFlags & ARO_LOCATION)
		{
			LocationLength.parse<WORD>(parser);
			LocationLength2.parse<WORD>(parser);
			if (LocationLength2 && LocationLength2 + 1 == LocationLength)
			{
				Location.parse(parser, LocationLength2);
			}
		}

		if (OverrideFlags & ARO_BUSYSTATUS)
		{
			BusyStatus.parse<DWORD>(parser);
		}

		if (OverrideFlags & ARO_ATTACHMENT)
		{
			Attachment.parse<DWORD>(parser);
		}

		if (OverrideFlags & ARO_SUBTYPE)
		{
			SubType.parse<DWORD>(parser);
		}

		if (OverrideFlags & ARO_APPTCOLOR)
		{
			AppointmentColor.parse<DWORD>(parser);
		}
	}

	ExtendedException::ExtendedException(std::shared_ptr<binaryParser>& parser, DWORD writerVersion2, WORD flags)
	{
		if (writerVersion2 >= 0x0003009)
		{
			ChangeHighlight.ChangeHighlightSize.parse<DWORD>(parser);
			ChangeHighlight.ChangeHighlightValue.parse<DWORD>(parser);
			if (ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
			{
				ChangeHighlight.Reserved.parse(parser, ChangeHighlight.ChangeHighlightSize - sizeof(DWORD), _MaxBytes);
			}
		}

		ReservedBlockEE1Size.parse<DWORD>(parser);
		ReservedBlockEE1.parse(parser, ReservedBlockEE1Size, _MaxBytes);

		if (flags & ARO_SUBJECT || flags & ARO_LOCATION)
		{
			StartDateTime.parse<DWORD>(parser);
			EndDateTime.parse<DWORD>(parser);
			OriginalStartDate.parse<DWORD>(parser);
		}

		if (flags & ARO_SUBJECT)
		{
			WideCharSubjectLength.parse<WORD>(parser);
			if (WideCharSubjectLength)
			{
				WideCharSubject.parse(parser, WideCharSubjectLength);
			}
		}

		if (flags & ARO_LOCATION)
		{
			WideCharLocationLength.parse<WORD>(parser);
			if (WideCharLocationLength)
			{
				WideCharLocation.parse(parser, WideCharLocationLength);
			}
		}

		if (flags & ARO_SUBJECT || flags & ARO_LOCATION)
		{
			ReservedBlockEE2Size.parse<DWORD>(parser);
			ReservedBlockEE2.parse(parser, ReservedBlockEE2Size, _MaxBytes);
		}
	}

	void AppointmentRecurrencePattern::Parse()
	{
		m_RecurrencePattern.parse(m_Parser, false);

		m_ReaderVersion2.parse<DWORD>(m_Parser);
		m_WriterVersion2.parse<DWORD>(m_Parser);
		m_StartTimeOffset.parse<DWORD>(m_Parser);
		m_EndTimeOffset.parse<DWORD>(m_Parser);
		m_ExceptionCount.parse<WORD>(m_Parser);

		if (m_ExceptionCount && m_ExceptionCount == m_RecurrencePattern.m_ModifiedInstanceCount &&
			m_ExceptionCount < _MaxEntriesSmall)
		{
			m_ExceptionInfo.reserve(m_ExceptionCount);
			for (WORD i = 0; i < m_ExceptionCount; i++)
			{
				m_ExceptionInfo.emplace_back(std::make_shared<ExceptionInfo>(m_Parser));
			}
		}

		m_ReservedBlock1Size.parse<DWORD>(m_Parser);
		m_ReservedBlock1.parse(m_Parser, m_ReservedBlock1Size, _MaxBytes);

		if (m_ExceptionCount && m_ExceptionCount == m_RecurrencePattern.m_ModifiedInstanceCount &&
			m_ExceptionCount < _MaxEntriesSmall && !m_ExceptionInfo.empty())
		{
			for (WORD i = 0; i < m_ExceptionCount; i++)
			{
				m_ExtendedException.emplace_back(
					std::make_shared<ExtendedException>(m_Parser, m_WriterVersion2, m_ExceptionInfo[i]->OverrideFlags));
			}
		}

		m_ReservedBlock2Size.parse<DWORD>(m_Parser);
		m_ReservedBlock2.parse(m_Parser, m_ReservedBlock2Size, _MaxBytes);
	}

	void AppointmentRecurrencePattern::ParseBlocks()
	{
		setRoot(m_RecurrencePattern.getBlock());

		terminateBlock();
		auto arpBlock = block{};
		arpBlock.setText(L"Appointment Recurrence Pattern: \r\n");
		arpBlock.addBlock(m_ReaderVersion2, L"ReaderVersion2: 0x%1!08X!\r\n", m_ReaderVersion2.getData());
		arpBlock.addBlock(m_WriterVersion2, L"WriterVersion2: 0x%1!08X!\r\n", m_WriterVersion2.getData());
		arpBlock.addBlock(
			m_StartTimeOffset,
			L"StartTimeOffset: 0x%1!08X! = %1!d! = %2!ws!\r\n",
			m_StartTimeOffset.getData(),
			RTimeToString(m_StartTimeOffset).c_str());
		arpBlock.addBlock(
			m_EndTimeOffset,
			L"EndTimeOffset: 0x%1!08X! = %1!d! = %2!ws!\r\n",
			m_EndTimeOffset.getData(),
			RTimeToString(m_EndTimeOffset).c_str());

		auto& exceptions = m_ExceptionCount;
		exceptions.setText(L"ExceptionCount: 0x%1!04X!\r\n", m_ExceptionCount.getData());

		if (!m_ExceptionInfo.empty())
		{
			auto i = 0;
			for (const auto& info : m_ExceptionInfo)
			{
				auto exception = block{};
				exception.addBlock(
					info->StartDateTime,
					L"ExceptionInfo[%1!d!].StartDateTime: 0x%2!08X! = %3!ws!\r\n",
					i,
					info->StartDateTime.getData(),
					RTimeToString(info->StartDateTime).c_str());
				exception.addBlock(
					info->EndDateTime,
					L"ExceptionInfo[%1!d!].EndDateTime: 0x%2!08X! = %3!ws!\r\n",
					i,
					info->EndDateTime.getData(),
					RTimeToString(info->EndDateTime).c_str());
				exception.addBlock(
					info->OriginalStartDate,
					L"ExceptionInfo[%1!d!].OriginalStartDate: 0x%2!08X! = %3!ws!\r\n",
					i,
					info->OriginalStartDate.getData(),
					RTimeToString(info->OriginalStartDate).c_str());
				auto szOverrideFlags = flags::InterpretFlags(flagOverrideFlags, info->OverrideFlags);
				exception.addBlock(
					info->OverrideFlags,
					L"ExceptionInfo[%1!d!].OverrideFlags: 0x%2!04X! = %3!ws!\r\n",
					i,
					info->OverrideFlags.getData(),
					szOverrideFlags.c_str());

				if (info->OverrideFlags & ARO_SUBJECT)
				{
					exception.addBlock(
						info->SubjectLength,
						L"ExceptionInfo[%1!d!].SubjectLength: 0x%2!04X! = %2!d!\r\n",
						i,
						info->SubjectLength.getData());
					exception.addBlock(
						info->SubjectLength2,
						L"ExceptionInfo[%1!d!].SubjectLength2: 0x%2!04X! = %2!d!\r\n",
						i,
						info->SubjectLength2.getData());

					exception.addBlock(
						info->Subject, L"ExceptionInfo[%1!d!].Subject: \"%2!hs!\"\r\n", i, info->Subject.c_str());
				}

				if (info->OverrideFlags & ARO_MEETINGTYPE)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						info->MeetingType, dispidApptStateFlags, const_cast<LPGUID>(&guid::PSETID_Appointment));
					exception.addBlock(
						info->MeetingType,
						L"ExceptionInfo[%1!d!].MeetingType: 0x%2!08X! = %3!ws!\r\n",
						i,
						info->MeetingType.getData(),
						szFlags.c_str());
				}

				if (info->OverrideFlags & ARO_REMINDERDELTA)
				{
					exception.addBlock(
						info->ReminderDelta,
						L"ExceptionInfo[%1!d!].ReminderDelta: 0x%2!08X!\r\n",
						i,
						info->ReminderDelta.getData());
				}

				if (info->OverrideFlags & ARO_REMINDER)
				{
					exception.addBlock(
						info->ReminderSet,
						L"ExceptionInfo[%1!d!].ReminderSet: 0x%2!08X!\r\n",
						i,
						info->ReminderSet.getData());
				}

				if (info->OverrideFlags & ARO_LOCATION)
				{
					exception.addBlock(
						info->LocationLength,
						L"ExceptionInfo[%1!d!].LocationLength: 0x%2!04X! = %2!d!\r\n",
						i,
						info->LocationLength.getData());
					exception.addBlock(
						info->LocationLength2,
						L"ExceptionInfo[%1!d!].LocationLength2: 0x%2!04X! = %2!d!\r\n",
						i,
						info->LocationLength2.getData());
					exception.addBlock(
						info->Location, L"ExceptionInfo[%1!d!].Location: \"%2!hs!\"\r\n", i, info->Location.c_str());
				}

				if (info->OverrideFlags & ARO_BUSYSTATUS)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						info->BusyStatus, dispidBusyStatus, const_cast<LPGUID>(&guid::PSETID_Appointment));
					exception.addBlock(
						info->BusyStatus,
						L"ExceptionInfo[%1!d!].BusyStatus: 0x%2!08X! = %3!ws!\r\n",
						i,
						info->BusyStatus.getData(),
						szFlags.c_str());
				}

				if (info->OverrideFlags & ARO_ATTACHMENT)
				{
					exception.addBlock(
						info->Attachment,
						L"ExceptionInfo[%1!d!].Attachment: 0x%2!08X!\r\n",
						i,
						info->Attachment.getData());
				}

				if (info->OverrideFlags & ARO_SUBTYPE)
				{
					exception.addBlock(
						info->SubType, L"ExceptionInfo[%1!d!].SubType: 0x%2!08X!\r\n", i, info->SubType.getData());
				}
				if (info->OverrideFlags & ARO_APPTCOLOR)
				{
					exception.addBlock(
						info->AppointmentColor,
						L"ExceptionInfo[%1!d!].AppointmentColor: 0x%2!08X!\r\n",
						i,
						info->AppointmentColor.getData());
				}

				exceptions.addBlock(exception);
				i++;
			}
		}

		arpBlock.addBlock(exceptions);
		auto& reservedBlock1 = m_ReservedBlock1Size;
		reservedBlock1.setText(L"ReservedBlock1Size: 0x%1!08X!", m_ReservedBlock1Size.getData());
		if (m_ReservedBlock1Size)
		{
			reservedBlock1.terminateBlock();
			reservedBlock1.addBlock(m_ReservedBlock1);
		}

		reservedBlock1.terminateBlock();
		arpBlock.addBlock(reservedBlock1);

		if (!m_ExtendedException.empty())
		{
			auto i = UINT{};
			for (const auto& ee : m_ExtendedException)
			{
				auto exception = block{};
				if (m_WriterVersion2 >= 0x00003009)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						ee->ChangeHighlight.ChangeHighlightValue,
						dispidChangeHighlight,
						const_cast<LPGUID>(&guid::PSETID_Appointment));
					exception.addBlock(
						ee->ChangeHighlight.ChangeHighlightSize,
						L"ExtendedException[%1!d!].ChangeHighlight.ChangeHighlightSize: 0x%2!08X!\r\n",
						i,
						ee->ChangeHighlight.ChangeHighlightSize.getData());
					exception.addBlock(
						ee->ChangeHighlight.ChangeHighlightValue,
						L"ExtendedException[%1!d!].ChangeHighlight.ChangeHighlightValue: 0x%2!08X! = %3!ws!\r\n",
						i,
						ee->ChangeHighlight.ChangeHighlightValue.getData(),
						szFlags.c_str());

					if (ee->ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
					{
						exception.addHeader(L"ExtendedException[%1!d!].ChangeHighlight.Reserved:", i);

						exception.addBlock(ee->ChangeHighlight.Reserved);
					}
				}

				exception.addBlock(
					ee->ReservedBlockEE1Size,
					L"ExtendedException[%1!d!].ReservedBlockEE1Size: 0x%2!08X!\r\n",
					i,
					ee->ReservedBlockEE1Size.getData());
				if (!ee->ReservedBlockEE1.empty())
				{
					exception.addBlock(ee->ReservedBlockEE1);
				}

				if (i < m_ExceptionInfo.size())
				{
					if (m_ExceptionInfo[i]->OverrideFlags & ARO_SUBJECT ||
						m_ExceptionInfo[i]->OverrideFlags & ARO_LOCATION)
					{
						exception.addBlock(
							ee->StartDateTime,
							L"ExtendedException[%1!d!].StartDateTime: 0x%2!08X! = %3\r\n",
							i,
							ee->StartDateTime.getData(),
							RTimeToString(ee->StartDateTime).c_str());
						exception.addBlock(
							ee->EndDateTime,
							L"ExtendedException[%1!d!].EndDateTime: 0x%2!08X! = %3!ws!\r\n",
							i,
							ee->EndDateTime.getData(),
							RTimeToString(ee->EndDateTime).c_str());
						exception.addBlock(
							ee->OriginalStartDate,
							L"ExtendedException[%1!d!].OriginalStartDate: 0x%2!08X! = %3!ws!\r\n",
							i,
							ee->OriginalStartDate.getData(),
							RTimeToString(ee->OriginalStartDate).c_str());
					}

					if (m_ExceptionInfo[i]->OverrideFlags & ARO_SUBJECT)
					{
						exception.addBlock(
							ee->WideCharSubjectLength,
							L"ExtendedException[%1!d!].WideCharSubjectLength: 0x%2!08X! = %2!d!\r\n",
							i,
							ee->WideCharSubjectLength.getData());
						exception.addBlock(
							ee->WideCharSubject,
							L"ExtendedException[%1!d!].WideCharSubject: \"%2!ws!\"\r\n",
							i,
							ee->WideCharSubject.c_str());
					}

					if (m_ExceptionInfo[i]->OverrideFlags & ARO_LOCATION)
					{
						exception.addBlock(
							ee->WideCharLocationLength,
							L"ExtendedException[%1!d!].WideCharLocationLength: 0x%2!08X! = %2!d!\r\n",
							i,
							ee->WideCharLocationLength.getData());
						exception.addBlock(
							ee->WideCharLocation,
							L"ExtendedException[%1!d!].WideCharLocation: \"%2!ws!\"\r\n",
							i,
							ee->WideCharLocation.c_str());
					}
				}

				exception.addBlock(
					ee->ReservedBlockEE2Size,
					L"ExtendedException[%1!d!].ReservedBlockEE2Size: 0x%2!08X!\r\n",
					i,
					ee->ReservedBlockEE2Size.getData());
				if (ee->ReservedBlockEE2Size)
				{
					exception.addBlock(ee->ReservedBlockEE2);
				}

				arpBlock.addBlock(exception);
				i++;
			}
		}

		auto& reservedBlock2 = m_ReservedBlock2Size;
		reservedBlock2.setText(L"ReservedBlock2Size: 0x%1!08X!", m_ReservedBlock2Size.getData());
		if (m_ReservedBlock2Size)
		{
			reservedBlock2.terminateBlock();
			reservedBlock2.addBlock(m_ReservedBlock2);
		}

		reservedBlock2.terminateBlock();
		arpBlock.addBlock(reservedBlock2);
		addBlock(arpBlock);
	}
} // namespace smartview