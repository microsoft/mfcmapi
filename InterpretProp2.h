#pragma once

// Function to convert property tags to their names
// Free lpszExactMatch and lpszPartialMatches with MAPIFreeBuffer
HRESULT PropTagToPropName(ULONG ulPropTag, BOOL bIsAB, LPTSTR* lpszExactMatch, LPTSTR* lpszPartialMatches);

HRESULT PropNameToPropTag(LPCTSTR lpszPropName, ULONG* ulPropTag);
HRESULT PropTypeNameToPropType(LPCTSTR lpszPropType, ULONG* ulPropType);
LPTSTR GUIDToStringAndName(LPCGUID lpGUID);
void GUIDNameToGUID(LPCTSTR szGUID, LPCGUID* lpGUID);

LPWSTR NameIDToPropName(LPMAPINAMEID lpNameID);

HRESULT InterpretFlags(const LPSPropValue lpProp, LPTSTR* szFlagString);
HRESULT InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, LPTSTR* szFlagString);
HRESULT InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, LPCTSTR szPrefix, LPTSTR* szFlagString);
CString AllFlagsToString(const ULONG ulFlagName,BOOL bHex);

// Uber property interpreter - given an LPSPropValue, produces all manner of strings
// All LPTSTR strings allocated with new, delete with delete[]
void InterpretProp(LPSPropValue lpProp, // optional property value
				   ULONG ulPropTag, // optional 'original' prop tag
				   LPMAPIPROP lpMAPIProp, // optional source object
				   LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
				   LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
				   BOOL bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
				   LPTSTR* lpszNameExactMatches, // Built from ulPropTag & bIsAB
				   LPTSTR* lpszNamePartialMatches, // Built from ulPropTag & bIsAB
				   CString* PropType, // Built from ulPropTag
				   CString* PropTag, // Built from ulPropTag
				   CString* PropString, // Built from lpProp
				   CString* AltPropString, // Built from lpProp
				   LPTSTR* lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
				   LPTSTR* lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
				   LPTSTR* lpszNamedPropDASL); // Built from ulPropTag & lpMAPIProp

HRESULT GetLargeBinaryProp(LPMAPIPROP lpMAPIProp, ULONG ulPropTag, LPSPropValue* lppProp);