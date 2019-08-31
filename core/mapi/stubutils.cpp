#include <core/stdafx.h>
#include <mapistub/library/mapiStubUtils.h>
#include <core/mapi/stubutils.h>

// Included for MFCMAPI tracing
#include <core/utility/output.h>

namespace mapistub
{
	void initStubCallbacks()
	{
		mapistub::logLoadMapiCallback = [](auto _1, auto _2) { output::DebugPrint(output::DBGLoadMAPI, _1, _2); };
		mapistub::logLoadLibraryCallback = [](auto _1, auto _2) { output::DebugPrint(output::DBGLoadLibrary, _1, _2); };
	}
} // namespace mapistub
