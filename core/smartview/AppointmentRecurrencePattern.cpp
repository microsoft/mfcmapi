#include <core/stdafx.h>
#include <core/smartview/AppointmentRecurrencePattern.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/guid.h>

namespace smartview
{
	ExceptionInfo::ExceptionInfo(const std::shared_ptr<binaryParser>& parser)
	{
		StartDateTime = blockT<DWORD>::parse(parser);
		EndDateTime = blockT<DWORD>::parse(parser);
		OriginalStartDate = blockT<DWORD>::parse(parser);
		OverrideFlags = blockT<WORD>::parse(parser);
		if (*OverrideFlags & ARO_SUBJECT)
		{
			SubjectLength = blockT<WORD>::parse(parser);
			SubjectLength2 = blockT<WORD>::parse(parser);
			if (*SubjectLength2 && *SubjectLength2 + 1 == *SubjectLength)
			{
				Subject = blockStringA::parse(parser, *SubjectLength2);
			}
		}

		if (*OverrideFlags & ARO_MEETINGTYPE)
		{
			MeetingType = blockT<DWORD>::parse(parser);
		}

		if (*OverrideFlags & ARO_REMINDERDELTA)
		{
			ReminderDelta = blockT<DWORD>::parse(parser);
		}

		if (*OverrideFlags & ARO_REMINDER)
		{
			ReminderSet = blockT<DWORD>::parse(parser);
		}

		if (*OverrideFlags & ARO_LOCATION)
		{
			LocationLength = blockT<WORD>::parse(parser);
			LocationLength2 = blockT<WORD>::parse(parser);
			if (*LocationLength2 && *LocationLength2 + 1 == *LocationLength)
			{
				Location = blockStringA::parse(parser, *LocationLength2);
			}
		}

		if (*OverrideFlags & ARO_BUSYSTATUS)
		{
			BusyStatus = blockT<DWORD>::parse(parser);
		}

		if (*OverrideFlags & ARO_ATTACHMENT)
		{
			Attachment = blockT<DWORD>::parse(parser);
		}

		if (*OverrideFlags & ARO_SUBTYPE)
		{
			SubType = blockT<DWORD>::parse(parser);
		}

		if (*OverrideFlags & ARO_APPTCOLOR)
		{
			AppointmentColor = blockT<DWORD>::parse(parser);
		}
	}

	ExtendedException::ExtendedException(const std::shared_ptr<binaryParser>& parser, DWORD writerVersion2, WORD flags)
	{
		if (writerVersion2 >= 0x0003009)
		{
			ChangeHighlight.ChangeHighlightSize = blockT<DWORD>::parse(parser);
			ChangeHighlight.ChangeHighlightValue = blockT<DWORD>::parse(parser);
			if (*ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
			{
				ChangeHighlight.Reserved =
					blockBytes::parse(parser, *ChangeHighlight.ChangeHighlightSize - sizeof(DWORD), _MaxBytes);
			}
		}

		ReservedBlockEE1Size = blockT<DWORD>::parse(parser);
		ReservedBlockEE1 = blockBytes::parse(parser, *ReservedBlockEE1Size, _MaxBytes);

		if (flags & ARO_SUBJECT || flags & ARO_LOCATION)
		{
			StartDateTime = blockT<DWORD>::parse(parser);
			EndDateTime = blockT<DWORD>::parse(parser);
			OriginalStartDate = blockT<DWORD>::parse(parser);
		}

		if (flags & ARO_SUBJECT)
		{
			WideCharSubjectLength = blockT<WORD>::parse(parser);
			if (*WideCharSubjectLength)
			{
				WideCharSubject = blockStringW::parse(parser, *WideCharSubjectLength);
			}
		}

		if (flags & ARO_LOCATION)
		{
			WideCharLocationLength = blockT<WORD>::parse(parser);
			if (*WideCharLocationLength)
			{
				WideCharLocation = blockStringW::parse(parser, *WideCharLocationLength);
			}
		}

		if (flags & ARO_SUBJECT || flags & ARO_LOCATION)
		{
			ReservedBlockEE2Size = blockT<DWORD>::parse(parser);
			ReservedBlockEE2 = blockBytes::parse(parser, *ReservedBlockEE2Size, _MaxBytes);
		}
	}

	void AppointmentRecurrencePattern::parse()
	{
		m_RecurrencePattern = block::parse<RecurrencePattern>(parser, false);

		m_ReaderVersion2 = blockT<DWORD>::parse(parser);
		m_WriterVersion2 = blockT<DWORD>::parse(parser);
		m_StartTimeOffset = blockT<DWORD>::parse(parser);
		m_EndTimeOffset = blockT<DWORD>::parse(parser);
		m_ExceptionCount = blockT<WORD>::parse(parser);

		if (*m_ExceptionCount && *m_ExceptionCount == *m_RecurrencePattern->m_ModifiedInstanceCount &&
			*m_ExceptionCount < _MaxEntriesSmall)
		{
			m_ExceptionInfo.reserve(*m_ExceptionCount);
			for (WORD i = 0; i < *m_ExceptionCount; i++)
			{
				m_ExceptionInfo.emplace_back(std::make_shared<ExceptionInfo>(parser));
			}
		}

		m_ReservedBlock1Size = blockT<DWORD>::parse(parser);
		m_ReservedBlock1 = blockBytes::parse(parser, *m_ReservedBlock1Size, _MaxBytes);

		if (*m_ExceptionCount && *m_ExceptionCount == *m_RecurrencePattern->m_ModifiedInstanceCount &&
			*m_ExceptionCount < _MaxEntriesSmall && !m_ExceptionInfo.empty())
		{
			for (WORD i = 0; i < *m_ExceptionCount; i++)
			{
				m_ExtendedException.emplace_back(
					std::make_shared<ExtendedException>(parser, *m_WriterVersion2, *m_ExceptionInfo[i]->OverrideFlags));
			}
		}

		m_ReservedBlock2Size = blockT<DWORD>::parse(parser);
		m_ReservedBlock2 = blockBytes::parse(parser, *m_ReservedBlock2Size, _MaxBytes);
	}

	void AppointmentRecurrencePattern::parseBlocks()
	{
		addChild(m_RecurrencePattern);

		auto arpBlock = create(L"Appointment Recurrence Pattern");
		addChild(arpBlock);

		arpBlock->addChild(m_ReaderVersion2, L"ReaderVersion2: 0x%1!08X!", m_ReaderVersion2->getData());
		arpBlock->addChild(m_WriterVersion2, L"WriterVersion2: 0x%1!08X!", m_WriterVersion2->getData());
		arpBlock->addChild(
			m_StartTimeOffset,
			L"StartTimeOffset: 0x%1!08X! = %1!d! = %2!ws!",
			m_StartTimeOffset->getData(),
			RTimeToString(*m_StartTimeOffset).c_str());
		arpBlock->addChild(
			m_EndTimeOffset,
			L"EndTimeOffset: 0x%1!08X! = %1!d! = %2!ws!",
			m_EndTimeOffset->getData(),
			RTimeToString(*m_EndTimeOffset).c_str());

		arpBlock->addChild(m_ExceptionCount);
		m_ExceptionCount->setText(L"ExceptionCount: 0x%1!04X!", m_ExceptionCount->getData());

		if (!m_ExceptionInfo.empty())
		{
			auto i = 0;
			for (const auto& info : m_ExceptionInfo)
			{
				auto exception = create(L"ExceptionInfo[%1!d!]", i);
				m_ExceptionCount->addChild(exception);

				exception->addChild(
					info->StartDateTime,
					L"StartDateTime: 0x%1!08X! = %2!ws!",
					info->StartDateTime->getData(),
					RTimeToString(*info->StartDateTime).c_str());
				exception->addChild(
					info->EndDateTime,
					L"EndDateTime: 0x%1!08X! = %2!ws!",
					info->EndDateTime->getData(),
					RTimeToString(*info->EndDateTime).c_str());
				exception->addChild(
					info->OriginalStartDate,
					L"OriginalStartDate: 0x%1!08X! = %2!ws!",
					info->OriginalStartDate->getData(),
					RTimeToString(*info->OriginalStartDate).c_str());
				auto szOverrideFlags = flags::InterpretFlags(flagOverrideFlags, *info->OverrideFlags);
				exception->addChild(
					info->OverrideFlags,
					L"OverrideFlags: 0x%1!04X! = %2!ws!",
					info->OverrideFlags->getData(),
					szOverrideFlags.c_str());

				if (*info->OverrideFlags & ARO_SUBJECT)
				{
					exception->addChild(
						info->SubjectLength, L"SubjectLength: 0x%1!04X! = %1!d!", info->SubjectLength->getData());
					exception->addChild(
						info->SubjectLength2, L"SubjectLength2: 0x%1!04X! = %1!d!", info->SubjectLength2->getData());

					exception->addChild(info->Subject, L"Subject: \"%1!hs!\"", info->Subject->c_str());
				}

				if (*info->OverrideFlags & ARO_MEETINGTYPE)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						*info->MeetingType, dispidApptStateFlags, const_cast<LPGUID>(&guid::PSETID_Appointment));
					exception->addChild(
						info->MeetingType,
						L"MeetingType: 0x%1!08X! = %2!ws!",
						info->MeetingType->getData(),
						szFlags.c_str());
				}

				if (*info->OverrideFlags & ARO_REMINDERDELTA)
				{
					exception->addChild(
						info->ReminderDelta, L"ReminderDelta: 0x%1!08X!", info->ReminderDelta->getData());
				}

				if (*info->OverrideFlags & ARO_REMINDER)
				{
					exception->addChild(info->ReminderSet, L"ReminderSet: 0x%1!08X!", info->ReminderSet->getData());
				}

				if (*info->OverrideFlags & ARO_LOCATION)
				{
					exception->addChild(
						info->LocationLength, L"LocationLength: 0x%1!04X! = %1!d!", info->LocationLength->getData());
					exception->addChild(
						info->LocationLength2, L"LocationLength2: 0x%1!04X! = %1!d!", info->LocationLength2->getData());
					exception->addChild(info->Location, L"Location: \"%1!hs!\"", info->Location->c_str());
				}

				if (*info->OverrideFlags & ARO_BUSYSTATUS)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						*info->BusyStatus, dispidBusyStatus, const_cast<LPGUID>(&guid::PSETID_Appointment));
					exception->addChild(
						info->BusyStatus,
						L"BusyStatus: 0x%1!08X! = %2!ws!",
						info->BusyStatus->getData(),
						szFlags.c_str());
				}

				if (*info->OverrideFlags & ARO_ATTACHMENT)
				{
					exception->addChild(info->Attachment, L"Attachment: 0x%1!08X!", info->Attachment->getData());
				}

				if (*info->OverrideFlags & ARO_SUBTYPE)
				{
					exception->addChild(info->SubType, L"SubType: 0x%1!08X!", info->SubType->getData());
				}

				if (*info->OverrideFlags & ARO_APPTCOLOR)
				{
					exception->addChild(
						info->AppointmentColor, L"AppointmentColor: 0x%1!08X!", info->AppointmentColor->getData());
				}

				i++;
			}
		}

		arpBlock->addChild(m_ReservedBlock1Size);
		m_ReservedBlock1Size->setText(L"ReservedBlock1Size: 0x%1!08X!", m_ReservedBlock1Size->getData());
		if (!m_ReservedBlock1->empty())
		{
			m_ReservedBlock1Size->addChild(m_ReservedBlock1);
		}

		if (!m_ExtendedException.empty())
		{
			auto i = UINT{};
			for (const auto& ee : m_ExtendedException)
			{
				auto exception = create(L"ExtendedException[%1!d!]", i);
				arpBlock->addChild(exception);

				if (*m_WriterVersion2 >= 0x00003009)
				{
					auto szFlags = InterpretNumberAsStringNamedProp(
						*ee->ChangeHighlight.ChangeHighlightValue,
						dispidChangeHighlight,
						const_cast<LPGUID>(&guid::PSETID_Appointment));
					exception->addChild(
						ee->ChangeHighlight.ChangeHighlightSize,
						L"ChangeHighlight.ChangeHighlightSize: 0x%1!08X!",
						ee->ChangeHighlight.ChangeHighlightSize->getData());
					exception->addChild(
						ee->ChangeHighlight.ChangeHighlightValue,
						L"ChangeHighlight.ChangeHighlightValue: 0x%1!08X! = %2!ws!",
						ee->ChangeHighlight.ChangeHighlightValue->getData(),
						szFlags.c_str());

					if (*ee->ChangeHighlight.ChangeHighlightSize > sizeof(DWORD))
					{
						exception->addHeader(L"ChangeHighlight.Reserved:");
						exception->addChild(ee->ChangeHighlight.Reserved);
					}
				}

				exception->addChild(
					ee->ReservedBlockEE1Size, L"ReservedBlockEE1Size: 0x%1!08X!", ee->ReservedBlockEE1Size->getData());
				if (!ee->ReservedBlockEE1->empty())
				{
					exception->addChild(ee->ReservedBlockEE1);
				}

				if (i < m_ExceptionInfo.size())
				{
					if (*m_ExceptionInfo[i]->OverrideFlags & ARO_SUBJECT ||
						*m_ExceptionInfo[i]->OverrideFlags & ARO_LOCATION)
					{
						exception->addChild(
							ee->StartDateTime,
							L"StartDateTime: 0x%1!08X! = %2",
							ee->StartDateTime->getData(),
							RTimeToString(*ee->StartDateTime).c_str());
						exception->addChild(
							ee->EndDateTime,
							L"EndDateTime: 0x%1!08X! = %2!ws!",
							ee->EndDateTime->getData(),
							RTimeToString(*ee->EndDateTime).c_str());
						exception->addChild(
							ee->OriginalStartDate,
							L"OriginalStartDate: 0x%1!08X! = %2!ws!",
							ee->OriginalStartDate->getData(),
							RTimeToString(*ee->OriginalStartDate).c_str());
					}

					if (*m_ExceptionInfo[i]->OverrideFlags & ARO_SUBJECT)
					{
						exception->addChild(
							ee->WideCharSubjectLength,
							L"WideCharSubjectLength: 0x%1!08X! = %1!d!",
							ee->WideCharSubjectLength->getData());
						exception->addChild(
							ee->WideCharSubject, L"WideCharSubject: \"%1!ws!\"", ee->WideCharSubject->c_str());
					}

					if (*m_ExceptionInfo[i]->OverrideFlags & ARO_LOCATION)
					{
						exception->addChild(
							ee->WideCharLocationLength,
							L"WideCharLocationLength: 0x%1!08X! = %1!d!",
							ee->WideCharLocationLength->getData());
						exception->addChild(
							ee->WideCharLocation, L"WideCharLocation: \"%1!ws!\"", ee->WideCharLocation->c_str());
					}
				}

				exception->addChild(
					ee->ReservedBlockEE2Size, L"ReservedBlockEE2Size: 0x%1!08X!", ee->ReservedBlockEE2Size->getData());
				if (!ee->ReservedBlockEE2->empty())
				{
					exception->addChild(ee->ReservedBlockEE2);
				}

				i++;
			}
		}

		arpBlock->addChild(m_ReservedBlock2Size);
		m_ReservedBlock2Size->setText(L"ReservedBlock2Size: 0x%1!08X!", m_ReservedBlock2Size->getData());
		if (!m_ReservedBlock2->empty())
		{
			m_ReservedBlock2Size->addChild(m_ReservedBlock2);
		}
	}
} // namespace smartview