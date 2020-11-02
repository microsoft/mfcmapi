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
		while (FlagArray[ulCurEntry].ulFlagName == ulFlagName)
		{
			switch (FlagArray[ulCurEntry].ulFlagType)
			{
			case flagFLAG:
				if (FlagArray[ulCurEntry].lFlagValue & lTempValue)
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
				}
				break;
			case flagVALUE:
				if (FlagArray[ulCurEntry].lFlagValue == lTempValue)
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = 0;
				}
				break;
			case flagVALUEHIGHBYTES:
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 16 & 0xFFFF))
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 16);
				}
				break;
			case flagVALUE3RDBYTE:
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 8 & 0xFF))
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 8);
				}
				break;
			case flagVALUE4THBYTE:
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0xFF))
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
				}
				break;
			case flagVALUELOWERNIBBLE:
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0x0F))
				{
					flags.push_back(FlagArray[ulCurEntry].lpszName);
					lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
				}
				break;
			case flagCLEARBITS:
				// find any bits we need to clear
				const auto lClearedBits = FlagArray[ulCurEntry].lFlagValue & lTempValue;
				// report what we found
				if (lClearedBits != 0)
				{
					flags.push_back(strings::format(L"0x%X", lClearedBits)); // STRING_OK
						// clear the bits out
					lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
				}
				break;
			}

			ulCurEntry++;
		}

		if (lTempValue || flags.empty())
		{
			flags.push_back(strings::format(L"0x%X", lTempValue)); // STRING_OK
		}

		return strings::join(flags, L" | ");
	}

	// Returns a list of all known flags/values for a flag name.
	// For instance, for flagFuzzyLevel, would return:
	// 0x00000000 FL_FULLSTRING\r\n\
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
		while (FlagArray[ulCurEntry].ulFlagName == ulFlagName)
		{
			if (flagCLEARBITS != FlagArray[ulCurEntry].ulFlagType)
			{
				flags.push_back(strings::formatmessage(
					bHex ? IDS_FLAGTOSTRINGHEX : IDS_FLAGTOSTRINGDEC,
					FlagArray[ulCurEntry].lFlagValue,
					FlagArray[ulCurEntry].lpszName));
			}

			ulCurEntry++;
		}

		return strings::join(flags, L"\r\n");
	}
} // namespace flags
