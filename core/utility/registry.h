// Header file for the MFCMAPI application registry functions

#pragma once

namespace registry
{
	static const wchar_t* RKEY_ROOT = L"SOFTWARE\\Microsoft\\MFCMAPI";

	enum class regKeyType
	{
		dword,
		string
	};

	enum class regOptionType
	{
		check,
		string,
		stringHex,
		stringDec
	};

	class __RegKey
	{
	public:
		std::wstring szKeyName;
		regKeyType ulRegKeyType{};
		regOptionType ulRegOptType{};
		DWORD ulDefDWORD{};
		DWORD ulCurDWORD{};
		std::wstring szDefSTRING;
		std::wstring szCurSTRING;
		bool bRefresh{};
		UINT uiOptionsPrompt{};

		__RegKey() = default;
		__RegKey(const __RegKey&) = delete;
		virtual __RegKey& operator=(__RegKey const&) = delete;
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
			ulRegKeyType = regKeyType::dword;
			ulRegOptType = regOptionType::check;
			ulDefDWORD = _default;
			ulCurDWORD = _default;
			bRefresh = _refresh;
			uiOptionsPrompt = _uiOptionsPrompt;
		}

		operator bool() const noexcept { return ulCurDWORD != 0; }

		boolRegKey& operator=(const bool val) noexcept
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
			const regOptionType _ulRegOptType,
			const DWORD _default,
			const bool _refresh,
			const int _uiOptionsPrompt)
			: __RegKey{}
		{
			szKeyName = _szKeyName;
			ulRegKeyType = regKeyType::dword;
			ulRegOptType = _ulRegOptType;
			ulDefDWORD = _default;
			ulCurDWORD = _default;
			bRefresh = _refresh;
			uiOptionsPrompt = _uiOptionsPrompt;
		}

		operator DWORD() const noexcept { return ulCurDWORD; }

		dwordRegKey& operator=(const DWORD val) noexcept
		{
			ulCurDWORD = val;
			return *this;
		}

		dwordRegKey& operator|=(const DWORD val) noexcept
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
			ulRegKeyType = regKeyType::string;
			ulRegOptType = regOptionType::string;
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
		_NODISCARD std::wstring::reference operator[](const std::wstring::size_type _Off) noexcept
		{
			return szCurSTRING[_Off];
		}
	};

	extern std::vector<__RegKey*> RegKeys;

	void WriteToRegistry();
	void ReadFromRegistry();

	_Check_return_ HKEY CreateRootKey();

	DWORD ReadDWORDFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ DWORD dwDefaultVal = {});
	std::wstring
	ReadStringFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ const std::wstring& szDefault = {});
	std::vector<BYTE>
	ReadBinFromRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValue, _In_ const bool bSecure = false);

	void WriteDWORDToRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValueName, DWORD dwValue);
	void WriteStringToRegistry(_In_ HKEY hKey, _In_ const std::wstring& szValueName, _In_ const std::wstring& szValue);
	void WriteBinToRegistry(
		_In_ HKEY hKey,
		_In_ const std::wstring& szValueName,
		_In_ const std::vector<BYTE>& binValue_In_,
		_In_ const bool bSecure = false);
	inline void WriteBinToRegistry(
		_In_ HKEY hKey,
		_In_ const std::wstring& szValueName,
		_In_ size_t cb,
		_In_ const BYTE* lpbin,
		_In_ const bool bSecure = false)
	{
		WriteBinToRegistry(hKey, szValueName, std::vector<BYTE>(lpbin, lpbin + cb), bSecure);
	}
	void PushOptionsToStub();

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
	extern boolRegKey preferOlmapi32;
	extern boolRegKey uiDiag;
	extern boolRegKey displayAboutDialog;
	extern wstringRegKey propertyColumnOrder;
	extern dwordRegKey namedPropBatchSize;
} // namespace registry