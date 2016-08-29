// MAPIProfileFunctions.h : Stand alone MAPI profile functions
#pragma once

wstring LaunchProfileWizard(
	_In_ HWND hParentWnd,
	ULONG ulFlags,
	_In_ string szServiceNameToAdd);

_Check_return_ HRESULT HrMAPIProfileExists(
	_In_ LPPROFADMIN lpProfAdmin,
	_In_ string& lpszProfileName);

_Check_return_ HRESULT HrCreateProfile(
	_In_ string& lpszProfileName); // profile name

_Check_return_ HRESULT HrAddServiceToProfile(
	_In_ string lpszServiceName, // Service Name
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	ULONG ulFlags, // Flags for CreateMsgService
	ULONG cPropVals, // Count of properties for ConfigureMsgService
	_In_opt_ LPSPropValue lpPropVals, // Properties for ConfigureMsgService
	_In_ string lpszProfileName); // profile name

_Check_return_ HRESULT HrAddExchangeToProfile(
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	_In_ string& lpszServerName,
	_In_ string& lpszMailboxName,
	_In_ string lpszProfileName);

_Check_return_ HRESULT HrAddPSTToProfile(
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	bool bUnicodePST,
	_In_ wstring& lpszPSTPath, // PST name
	_In_ string lpszProfileName, // profile name
	bool bPasswordSet, // whether or not to include a password
	_In_ string& lpszPassword); // password to include

_Check_return_ HRESULT HrRemoveProfile(
	_In_ string lpszProfileName);

_Check_return_ HRESULT HrSetDefaultProfile(
	_In_ string lpszProfileName);

_Check_return_ HRESULT OpenProfileSection(_In_ LPSERVICEADMIN lpServiceAdmin, _In_ LPSBinary lpServiceUID, _Deref_out_opt_ LPPROFSECT* lppProfSect);

_Check_return_ HRESULT OpenProfileSection(_In_ LPPROVIDERADMIN lpProviderAdmin, _In_ LPSBinary lpProviderUID, _Deref_out_ LPPROFSECT* lppProfSect);

void AddServicesToMapiSvcInf();
void RemoveServicesFromMapiSvcInf();
void GetMAPISVCPath(_Inout_z_count_(cchMAPIDir) LPTSTR szMAPIDir, ULONG cchMAPIDir);
void DisplayMAPISVCPath(_In_ CWnd* pParentWnd);

// http://msdn2.microsoft.com/en-us/library/bb820969.aspx
struct EXCHANGE_STORE_VERSION_NUM
{
	WORD wMajorVersion;
	WORD wMinorVersion;
	WORD wBuild;
	WORD wMinorBuild;
};

_Check_return_ HRESULT GetProfileServiceVersion(
	_In_ string lpszProfileName,
	_Out_ ULONG* lpulServerVersion,
	_Out_ EXCHANGE_STORE_VERSION_NUM* lpStoreVersion,
	_Out_ bool* lpbFoundServerVersion,
	_Out_ bool* lpbFoundServerFullVersion);

_Check_return_ HRESULT HrCopyProfile(
	_In_ string lpszOldProfileName,
	_In_ string lpszNewProfileName);