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
		auto bNeedSeparator = false;

		auto lTempValue = lFlagValue;
		std::wstring szTempString;
		for (; FlagArray[ulCurEntry].ulFlagName == ulFlagName; ulCurEntry++)
		{
			if (flagFLAG == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue & lTempValue)
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
					bNeedSeparator = true;
				}
			}
			else if (flagVALUE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == lTempValue)
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = 0;
					bNeedSeparator = true;
				}
			}
			else if (flagVALUEHIGHBYTES == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 16 & 0xFFFF))
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 16);
					bNeedSeparator = true;
				}
			}
			else if (flagVALUE3RDBYTE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 8 & 0xFF))
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 8);
					bNeedSeparator = true;
				}
			}
			else if (flagVALUE4THBYTE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0xFF))
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
					bNeedSeparator = true;
				}
			}
			else if (flagVALUELOWERNIBBLE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0x0F))
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
					bNeedSeparator = true;
				}
			}
			else if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
			{
				// find any bits we need to clear
				const auto lClearedBits = FlagArray[ulCurEntry].lFlagValue & lTempValue;
				// report what we found
				if (0 != lClearedBits)
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += strings::format(L"0x%X", lClearedBits); // STRING_OK
						// clear the bits out
					lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
					bNeedSeparator = true;
				}
			}
		}

		// We know if we've found anything already because bNeedSeparator will be true
		// If bNeedSeparator isn't true, we found nothing and need to tack on
		// Otherwise, it's true, and we only tack if lTempValue still has something in it
		if (!bNeedSeparator || lTempValue)
		{
			if (bNeedSeparator)
			{
				szTempString += L" | "; // STRING_OK
			}

			szTempString += strings::format(L"0x%X", lTempValue); // STRING_OK
		}

		return szTempString;
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
		std::wstring szFlagString;
		if (!ulFlagName) return szFlagString;
		if (FlagArray.empty()) return szFlagString;

		ULONG ulCurEntry = 0;
		std::wstring szTempString;

		while (ulCurEntry < FlagArray.size() && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
		{
			ulCurEntry++;
		}

		if (ulCurEntry == FlagArray.size() || FlagArray[ulCurEntry].ulFlagName != ulFlagName) return szFlagString;

		// We've matched our flag name to the array - we SHOULD return a string at this point
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
					szFlagString += strings::formatmessage(
						IDS_FLAGTOSTRINGHEX, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName);
				}
				else
				{
					szFlagString += strings::formatmessage(
						IDS_FLAGTOSTRINGDEC, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName);
				}
			}
		}

		return szFlagString;
	}
} // namespace flags
