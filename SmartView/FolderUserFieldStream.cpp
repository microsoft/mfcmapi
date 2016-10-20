#include "stdafx.h"
#include "FolderUserFieldStream.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

FolderUserFieldStream::FolderUserFieldStream() : m_FolderUserFieldsAnsiCount(0), m_FolderUserFieldsUnicodeCount(0) {}

void FolderUserFieldStream::Parse()
{
	m_Parser.GetDWORD(&m_FolderUserFieldsAnsiCount);
	if (m_FolderUserFieldsAnsiCount && m_FolderUserFieldsAnsiCount < _MaxEntriesSmall)
	{
		for (DWORD i = 0; i < m_FolderUserFieldsAnsiCount; i++)
		{
			FolderFieldDefinitionA folderFieldDefinitionA;
			m_Parser.GetDWORD(&folderFieldDefinitionA.FieldType);
			m_Parser.GetWORD(&folderFieldDefinitionA.FieldNameLength);

			if (folderFieldDefinitionA.FieldNameLength &&
				folderFieldDefinitionA.FieldNameLength < _MaxEntriesSmall)
			{
				folderFieldDefinitionA.FieldName = m_Parser.GetStringA(folderFieldDefinitionA.FieldNameLength);
			}

			folderFieldDefinitionA.Common = BinToFolderFieldDefinitionCommon();
			m_FieldDefinitionsA.push_back(folderFieldDefinitionA);
		}
	}

	m_Parser.GetDWORD(&m_FolderUserFieldsUnicodeCount);
	if (m_FolderUserFieldsUnicodeCount && m_FolderUserFieldsUnicodeCount < _MaxEntriesSmall)
	{
		for (DWORD i = 0; i < m_FolderUserFieldsUnicodeCount; i++)
		{
			FolderFieldDefinitionW folderFieldDefinitionW;
			m_Parser.GetDWORD(&folderFieldDefinitionW.FieldType);
			m_Parser.GetWORD(&folderFieldDefinitionW.FieldNameLength);

			if (folderFieldDefinitionW.FieldNameLength &&
				folderFieldDefinitionW.FieldNameLength < _MaxEntriesSmall)
			{
				folderFieldDefinitionW.FieldName = m_Parser.GetStringW(folderFieldDefinitionW.FieldNameLength);
			}

			folderFieldDefinitionW.Common = BinToFolderFieldDefinitionCommon();
			m_FieldDefinitionsW.push_back(folderFieldDefinitionW);
		}
	}
}

FolderFieldDefinitionCommon FolderUserFieldStream::BinToFolderFieldDefinitionCommon()
{
	FolderFieldDefinitionCommon common;

	m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), reinterpret_cast<LPBYTE>(&common.PropSetGuid));
	m_Parser.GetDWORD(&common.fcapm);
	m_Parser.GetDWORD(&common.dwString);
	m_Parser.GetDWORD(&common.dwBitmap);
	m_Parser.GetDWORD(&common.dwDisplay);
	m_Parser.GetDWORD(&common.iFmt);
	m_Parser.GetWORD(&common.wszFormulaLength);
	if (common.wszFormulaLength &&
		common.wszFormulaLength < _MaxEntriesLarge)
	{
		common.wszFormula = m_Parser.GetStringW(common.wszFormulaLength);
	}

	return common;
}

_Check_return_ wstring FolderUserFieldStream::ToStringInternal()
{
	wstring szFolderUserFieldStream;
	wstring szTmp;

	szFolderUserFieldStream = formatmessage(
		IDS_FIELDHEADER,
		m_FolderUserFieldsAnsiCount);

	if (m_FolderUserFieldsAnsiCount && m_FieldDefinitionsA.size())
	{
		for (DWORD i = 0; i < m_FieldDefinitionsA.size(); i++)
		{
			auto szGUID = GUIDToString(&m_FieldDefinitionsA[i].Common.PropSetGuid);
			auto szFieldType = InterpretFlags(flagFolderType, m_FieldDefinitionsA[i].FieldType);
			auto szFieldcap = InterpretFlags(flagFieldCap, m_FieldDefinitionsA[i].Common.fcapm);

			szTmp = formatmessage(
				IDS_FIELDANSIFIELD,
				i,
				m_FieldDefinitionsA[i].FieldType, szFieldType.c_str(),
				m_FieldDefinitionsA[i].FieldNameLength,
				m_FieldDefinitionsA[i].FieldName.c_str(),
				szGUID.c_str(),
				m_FieldDefinitionsA[i].Common.fcapm, szFieldcap.c_str(),
				m_FieldDefinitionsA[i].Common.dwString,
				m_FieldDefinitionsA[i].Common.dwBitmap,
				m_FieldDefinitionsA[i].Common.dwDisplay,
				m_FieldDefinitionsA[i].Common.iFmt,
				m_FieldDefinitionsA[i].Common.wszFormulaLength,
				m_FieldDefinitionsA[i].Common.wszFormula.c_str());
			szFolderUserFieldStream += szTmp;
		}
	}

	szTmp = formatmessage(
		IDS_FIELDUNICODEHEADER,
		m_FolderUserFieldsUnicodeCount);
	szFolderUserFieldStream += szTmp;

	if (m_FolderUserFieldsUnicodeCount&& m_FieldDefinitionsW.size())
	{
		for (DWORD i = 0; i < m_FieldDefinitionsW.size(); i++)
		{
			auto szGUID = GUIDToString(&m_FieldDefinitionsW[i].Common.PropSetGuid);
			auto szFieldType = InterpretFlags(flagFolderType, m_FieldDefinitionsW[i].FieldType);
			auto szFieldcap = InterpretFlags(flagFieldCap, m_FieldDefinitionsW[i].Common.fcapm);

			szTmp = formatmessage(
				IDS_FIELDUNICODEFIELD,
				i,
				m_FieldDefinitionsW[i].FieldType, szFieldType.c_str(),
				m_FieldDefinitionsW[i].FieldNameLength,
				m_FieldDefinitionsW[i].FieldName.c_str(),
				szGUID.c_str(),
				m_FieldDefinitionsW[i].Common.fcapm, szFieldcap.c_str(),
				m_FieldDefinitionsW[i].Common.dwString,
				m_FieldDefinitionsW[i].Common.dwBitmap,
				m_FieldDefinitionsW[i].Common.dwDisplay,
				m_FieldDefinitionsW[i].Common.iFmt,
				m_FieldDefinitionsW[i].Common.wszFormulaLength,
				m_FieldDefinitionsW[i].Common.wszFormula.c_str());
			szFolderUserFieldStream += szTmp;
		}
	}

	return szFolderUserFieldStream;
}