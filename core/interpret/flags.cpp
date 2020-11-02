#include <core/stdafx.h>
#include <core/interpret/flags.h>
#include <core/utility/strings.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>

namespace flags
{
	// Interprets a flag value according to a flag name and returns a string
	// Will not return a string if the flag name is not recognized
	std::wstring InterpretFlags(ULONG ulFlagName, LONG lFlagValue)
	{
		ULONG ulCurEntry = 0;

		if (FlagArray.empty()) return L"";

		while (ulCurEntry < FlagArray.size() && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
		{
			ulCurEntry++;
		}

		// Don't run off the end of the array
		if (FlagArray.size() == ulCurEntry) return L"";
		if (FlagArray[ulCurEntry].ulFlagName != ulFlagName) return L"";

		// We've matched our flag name to the array - we SHOULD return a string at this point
		auto flags = std::vector<std::wstring>{};
		auto lTempValue = lFlagValue;
		for (; FlagArray[ulCurEntry].ulFlagName == ulFlagName; ulCurEntry++)
		{
			if (flagFLAG == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue & lTempValue)
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
				}
			}
			else if (flagVALUE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == lTempValue)
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = 0;
				}
			}
			else if (flagVALUEHIGHBYTES == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 16 & 0xFFFF))
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 16);
				}
			}
			else if (flagVALUE3RDBYTE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 8 & 0xFF))
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 8);
				}
			}
			else if (flagVALUE4THBYTE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0xFF))
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
				}
			}
			else if (flagVALUELOWERNIBBLE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0x0F))
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
				}
			}
			else if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
			{
				// find any bits we need to clear
				const auto lClearedBits = FlagArray[ulCurEntry].lFlagValue & lTempValue;
				// report what we found
				if (0 != lClearedBits)
				{
					flags.push_back(strings::format(L"0x%X", lClearedBits)); // STRING_OK
						// clear the bits out
					lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
				}
			}
		}

		if (lTempValue || flags.empty())
		{
			flags.push_back(strings::format(L"0x%X", lTempValue)); // STRING_OK
		}

		return strings::join(flags, L" | ");
	}

	// Returns a list of all known flags/values for a flag name.
	// For instance, for flagFuzzyLevel, would return:
	// \r\n0x00000000 FL_FULLSTRING\r\n\
	 // 0x00000001 FL_SUBSTRING\r\n\
	 // 0x00000002 FL_PREFIX\r\n\
	 // 0x00010000 FL_IGNORECASE\r\n\
	 // 0x00020000 FL_IGNORENONSPACE\r\n\
	 // 0x00040000 FL_LOOSE
	//
	// Since the string is always appended to a prompt we include \r\n at the start
	std::wstring AllFlagsToString(ULONG ulFlagName, bool bHex)
	{
		if (!ulFlagName) return L"";
		if (FlagArray.empty()) return L"";

		ULONG ulCurEntry = 0;

		while (ulCurEntry < FlagArray.size() && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
		{
			ulCurEntry++;
		}

		if (ulCurEntry == FlagArray.size() || FlagArray[ulCurEntry].ulFlagName != ulFlagName) return L"";

		// We've matched our flag name to the array - we SHOULD return a string at this point
		auto flags = std::vector<std::wstring>{};
		for (; FlagArray[ulCurEntry].ulFlagName == ulFlagName; ulCurEntry++)
		{
			if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
			{
				// keep going
			}
			else
			{
				if (bHex)
				{
					flags.push_back(strings::formatmessage(
						IDS_FLAGTOSTRINGHEX, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName));
				}
				else
				{
					flags.push_back(strings::formatmessage(
						IDS_FLAGTOSTRINGDEC, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName));
				}
			}
		}

		return strings::join(flags, L"");
	}
} // namespace flags
