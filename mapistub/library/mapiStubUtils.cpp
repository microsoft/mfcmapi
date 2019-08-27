#include <mapistub/library/mapiStubUtils.h>
#include <Windows.h>
#include <Msi.h>
#include <winreg.h>

#include <stdlib.h>

namespace mapistub
{
	// Sequence number which is incremented every time we set our MAPI handle which will
	// cause a re-fetch of all stored function pointers
	volatile ULONG g_ulDllSequenceNum = 1;
} // namespace mapistub
