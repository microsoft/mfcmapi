// Header file for the MFCMAPI application registry functions

#pragma once

namespace registry
{
	static const wchar_t* RKEY_ROOT = L"SOFTWARE\\Microsoft\\MFCMAPI";

	enum __REGKEYTYPES
	{
		regDWORD,
		regSTRING
	};

	enum __REGOPTIONTYPE
	{
		regoptCheck,
		regoptString,
		regoptStringHex,
		regoptStringDec
	};

	class __RegKey
	{
	public:
		std::wstring szKeyName;
		__REGKEYTYPES ulRegKeyType;
		__REGOPTIONTYPE ulRegOptType;
		DWORD ulDefDWORD;
		DWORD ulCurDWORD;
		std::wstring szDefSTRING;
		std::wstring szCurSTRING;
		bool bRefresh;
		UINT uiOptionsPrompt;

		__RegKey(const __RegKey&) = delete;
		__RegKey& operator=(__RegKey const&) = delete;
		__RegKey(__RegKey&&) = delete;
		__RegKey& operator=(__RegKey&&) = delete;
		~__RegKey() = default;
	};

	class boolRegKey : public __RegKey
	{
	public:
		boolRegKey(const std::wstring& _szKeyName, const bool _default, const bool _refresh, const int _uiOptionsPrompt)
			: __RegKey{}
		{
			szKeyName = _szKeyName;
			ulRegKeyType = regDWORD;
			ulRegOptType = regoptCheck;
			ulDefDWORD = _default;
			ulCurDWORD = _default;
			bRefresh = _refresh;
			uiOptionsPrompt = _uiOptionsPrompt;
		}

		operator bool() const { return ulCurDWORD != 0; }

		boolRegKey& operator=(const bool val)
		{
			ulCurDWORD = val;
			return *this;
		}
	};

	class dwordRegKey : public __RegKey
	{
	public:
		dwordRegKey(
			const std::wstring& _szKeyName,
			const __REGOPTIONTYPE _ulRegOptType,
			const DWORD _default,
			const bool _refresh,
			const int _uiOptionsPrompt)
			: __RegKey{}
		{
			szKeyName = _szKeyName;
			ulRegKeyType = regDWORD;
			ulRegOptType = _ulRegOptType;
			ulDefDWORD = _default;
			ulCurDWORD = _default;
			bRefresh = _refresh;
			uiOptionsPrompt = _uiOptionsPrompt;
		}

		operator DWORD() const { return ulCurDWORD; }

		dwordRegKey& operator=(const DWORD val)
		{
			ulCurDWORD = val;
			return *this;
		}

		dwordRegKey& operator|=(const DWORD val)
		{
			ulCurDWORD |= val;
			return *this;
		}
	};

	class wstringRegKey : public __RegKey
	{
	public:
		wstringRegKey(
			const std::wstring& _szKeyName,
			const std::wstring& _default,
			const bool _refresh,
			const int _uiOptionsPrompt)
			: __RegKey{}
		{
			szKeyName = _szKeyName;
			ulRegKeyType = regSTRING;
			ulRegOptType = regoptString;
			szDefSTRING = _default;
			szCurSTRING = _default;
			bRefresh = _refresh;
			uiOptionsPrompt = _uiOptionsPrompt;
		}

		operator std::wstring() const { return szCurSTRING; }

		wstringRegKey& operator=(std::wstring val)
		{
			szCurSTRING = std::move(val);
			return *this;
		}

		_NODISCARD std::wstring::size_type length() const noexcept { return szCurSTRING.length(); }
		_NODISCARD bool empty() const noexcept { return szCurSTRING.empty(); }
		void clear() noexcept { szCurSTRING.clear(); }
		void push_back(const wchar_t _Ch) { szCurSTRING.push_back(_Ch); }
		_NODISCARD std::wstring::reference operator[](const std::wstring::size_type _Off) { return szCurSTRING[_Off]; }
	};

	extern std::vector<__RegKey*> RegKeys;

	void WriteToRegistry();
	void ReadFromRegistry();

	_Check_return_ HKEY CreateRootKey();

	DWORD ReadDWORDFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ DWORD dwDefaultVal = 0);
	std::wstring ReadStringFromRegistry(
		_In_ HKEY hKey,
		_In_ const std::wstring& szValue,
		_In_ const std::wstring& szDefault = L"");

	void WriteStringToRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValueName, _In_ const std::wstring& szValue);

	extern dwordRegKey debugTag;
	extern boolRegKey debugToFile;
	extern wstringRegKey debugFileName;
	extern boolRegKey getPropNamesOnAllProps;
	extern boolRegKey parseNamedProps;
	extern dwordRegKey throttleLevel;
	extern boolRegKey hierExpandNotifications;
	extern boolRegKey hierRootNotifs;
	extern boolRegKey doSmartView;
	extern boolRegKey onlyAdditionalProperties;
	extern boolRegKey useRowDataForSinglePropList;
	extern boolRegKey useGetPropList;
	extern boolRegKey preferUnicodeProps;
	extern boolRegKey cacheNamedProps;
	extern boolRegKey allowDupeColumns;
	extern boolRegKey doColumnNames;
	extern boolRegKey editColumnsOnLoad;
	extern boolRegKey forceMDBOnline;
	extern boolRegKey forceMapiNoCache;
	extern boolRegKey allowPersistCache;
	extern boolRegKey useIMAPIProgress;
	extern boolRegKey useMessageRaw;
	extern boolRegKey suppressNotFound;
	extern boolRegKey heapEnableTerminationOnCorruption;
	extern boolRegKey loadAddIns;
	extern boolRegKey forceOutlookMAPI;
	extern boolRegKey forceSystemMAPI;
	extern boolRegKey hexDialogDiag;
	extern boolRegKey displayAboutDialog;
	extern wstringRegKey propertyColumnOrder;
} // namespace registry