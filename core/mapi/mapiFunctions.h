// Stand alone MAPI functions
#pragma once

namespace mapi
{
	// Safely cast across MAPI interfaces. Result is addrefed and must be released.
	template <class T> T safe_cast(IUnknown* src);
} // namespace mapi