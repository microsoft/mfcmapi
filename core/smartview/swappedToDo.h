#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockT.h>
#include <core/smartview/block/blockStringW.h>

namespace smartview
{
	// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxoflag/74057397-a90e-4454-bd48-3923e09c50b0
	class swappedToDo : public block
	{
	public:
		bool bBadData{};

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> ulVersion = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwFlags = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwToDoItem = emptyT<DWORD>();
		std::shared_ptr<blockStringW> wszFlagTo = emptySW();
		std::shared_ptr<blockT<DWORD>> rtmStartDate = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> rtmDueDate = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> rtmReminder = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> fReminderSet = emptyT<DWORD>();
	};
} // namespace smartview