#include <core/stdafx.h>
#include <core/smartview/swappedToDo.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>

namespace smartview
{
	void swappedToDo::parse()
	{
		ulVersion = blockT<DWORD>::parse(parser);
		dwFlags = blockT<DWORD>::parse(parser);
		dwToDoItem = blockT<DWORD>::parse(parser);
		wszFlagTo = blockStringW::parse(parser, 256); // 512 bytes is 256 characters
		rtmStartDate = blockT<DWORD>::parse(parser);
		rtmDueDate = blockT<DWORD>::parse(parser);
		rtmReminder = blockT<DWORD>::parse(parser);
		fReminderSet = blockT<DWORD>::parse(parser);
	}

	void swappedToDo::parseBlocks()
	{
		setText(L"Swapped ToDo Data");
		addChild(ulVersion, L"ulVersion = 0x%1!08X!", ulVersion->getData());
		const auto szFlags = flags::InterpretFlags(flagToDoSwapFlag, *dwFlags);
		addChild(dwFlags, L"dwFlags = 0x%1!02X! = %2!ws!", dwFlags->getData(), szFlags.c_str());
		addChild(dwToDoItem, L"dwToDoItem = 0x%1!08X!", dwToDoItem->getData());
		addChild(wszFlagTo, L"wszFlagTo = %1!ws!", wszFlagTo->c_str());
		addChild(
			rtmStartDate,
			L"rtmStartDate = 0x%1!08X! = %2!ws!",
			rtmStartDate->getData(),
			RTimeToString(*rtmStartDate).c_str());
		addChild(
			rtmDueDate, L"rtmDueDate = 0x%1!08X! = %2!ws!", rtmDueDate->getData(), RTimeToString(*rtmDueDate).c_str());
		addChild(
			rtmReminder,
			L"rtmReminder = 0x%1!08X! = %2!ws!",
			rtmReminder->getData(),
			RTimeToString(*rtmReminder).c_str());
		addChild(fReminderSet, L"fReminderSet: %1!ws!", *fReminderSet ? L"true" : L"false");
	}
} // namespace smartview