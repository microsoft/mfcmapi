#pragma once

// Function to convert property tags to their names
// Free lpszExactMatch and lpszPartialMatches with MAPIFreeBuffer
_Check_return_ HRESULT PropTagToPropName(ULONG ulPropTag, BOOL bIsAB, _Deref_opt_out_opt_z_ LPTSTR* lpszExactMatch, _Deref_opt_out_opt_z_ LPTSTR* lpszPartialMatches);

_Check_return_ HRESULT PropNameToPropTag(_In_z_ LPCTSTR lpszPropName, _Out_ ULONG* ulPropTag);
_Check_return_ ULONG PropTypeNameToPropTypeW(_In_z_ LPCWSTR lpszPropType);
_Check_return_ ULONG PropTypeNameToPropTypeA(_In_z_ LPCSTR lpszPropType);
#ifdef UNICODE
#define PropTypeNameToPropType PropTypeNameToPropTypeW
#else
#define PropTypeNameToPropType PropTypeNameToPropTypeA
#endif

_Check_return_ LPTSTR GUIDToStringAndName(_In_opt_ LPCGUID lpGUID);
void GUIDNameToGUID(_In_z_ LPCTSTR szGUID, _Deref_out_opt_ LPCGUID* lpGUID);

_Check_return_ LPWSTR NameIDToPropName(_In_ LPMAPINAMEID lpNameID);

void InterpretFlags(_In_ const LPSPropValue lpProp, _Deref_out_opt_z_ LPTSTR* szFlagString);
void InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, _Deref_out_opt_z_ LPTSTR* szFlagString);
void InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, _In_z_ LPCTSTR szPrefix, _Deref_out_opt_z_ LPTSTR* szFlagString);
_Check_return_ CString AllFlagsToString(const ULONG ulFlagName, BOOL bHex);

// Uber property interpreter - given an LPSPropValue, produces all manner of strings
// All LPTSTR strings allocated with new, delete with delete[]
void InterpretProp(_In_opt_ LPSPropValue lpProp, // optional property value
				   ULONG ulPropTag, // optional 'original' prop tag
				   _In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
				   _In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
				   _In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
				   BOOL bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
				   _Deref_out_opt_z_ LPTSTR* lpszNameExactMatches, // Built from ulPropTag & bIsAB
				   _Deref_out_opt_z_ LPTSTR* lpszNamePartialMatches, // Built from ulPropTag & bIsAB
				   _In_opt_ CString* PropType, // Built from ulPropTag
				   _In_opt_ CString* PropTag, // Built from ulPropTag
				   _In_opt_ CString* PropString, // Built from lpProp
				   _In_opt_ CString* AltPropString, // Built from lpProp
				   _Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
				   _Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
				   _Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropDASL); // Built from ulPropTag & lpMAPIProp

_Check_return_ HRESULT GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);