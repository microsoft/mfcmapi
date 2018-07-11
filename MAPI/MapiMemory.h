#pragma once

namespace mapi
{
	// Allocates an object using MAPI's memory allocators.
	// Switches allocators based on whether parent is passed.
	// Zero's memory before returning.
	template <typename T> T allocate(ULONG size, _In_opt_ LPVOID parent = nullptr)
	{
		LPVOID obj = nullptr;

		if (parent)
		{
			EC_H_S(MAPIAllocateMore(size, parent, &obj));
		}
		else
		{
			EC_H_S(MAPIAllocateBuffer(size, &obj));
		}

		if (obj)
		{
			memset(obj, 0, size);
		}

		return static_cast<T>(obj);
	}
}