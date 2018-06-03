#include <StdAfx.h>
#include <Interpret/SmartView/FolderUserFieldStream.h>
#include <Interpret/Guids.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	FolderUserFieldStream::FolderUserFieldStream() : m_FolderUserFieldsAnsiCount(0), m_FolderUserFieldsUnicodeCount(0) {}

	void FolderUserFieldStream::Parse()
	{
		m_FolderUserFieldsAnsiCount = m_Parser.Get<DWORD>();
		if (m_FolderUserFieldsAnsiCount && m_FolderUserFieldsAnsiCount < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_FolderUserFieldsAnsiCount; i++)
			{
				FolderFieldDefinitionA folderFieldDefinitionA;
				folderFieldDefinitionA.FieldType = m_Parser.Get<DWORD>();
				folderFieldDefinitionA.FieldNameLength = m_Parser.Get<WORD>();

				if (folderFieldDefinitionA.FieldNameLength &&
					folderFieldDefinitionA.FieldNameLength < _MaxEntriesSmall)
				{
					folderFieldDefinitionA.FieldName = m_Parser.GetStringA(folderFieldDefinitionA.FieldNameLength);
				}

				folderFieldDefinitionA.Common = BinToFolderFieldDefinitionCommon();
				m_FieldDefinitionsA.push_back(folderFieldDefinitionA);
			}
		}

		m_FolderUserFieldsUnicodeCount = m_Parser.Get<DWORD>();
		if (m_FolderUserFieldsUnicodeCount && m_FolderUserFieldsUnicodeCount < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_FolderUserFieldsUnicodeCount; i++)
			{
				FolderFieldDefinitionW folderFieldDefinitionW;
				folderFieldDefinitionW.FieldType = m_Parser.Get<DWORD>();
				folderFieldDefinitionW.FieldNameLength = m_Parser.Get<WORD>();

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

		common.PropSetGuid = m_Parser.Get<GUID>();
		common.fcapm = m_Parser.Get<DWORD>();
		common.dwString = m_Parser.Get<DWORD>();
		common.dwBitmap = m_Parser.Get<DWORD>();
		common.dwDisplay = m_Parser.Get<DWORD>();
		common.iFmt = m_Parser.Get<DWORD>();
		common.wszFormulaLength = m_Parser.Get<WORD>();
		if (common.wszFormulaLength &&
			common.wszFormulaLength < _MaxEntriesLarge)
		{
			common.wszFormula = m_Parser.GetStringW(common.wszFormulaLength);
		}

		return common;
	}

	_Check_return_ std::wstring FolderUserFieldStream::ToStringInternal()
	{
		std::wstring szTmp;

		auto szFolderUserFieldStream = strings::formatmessage(
			IDS_FIELDHEADER,
			m_FolderUserFieldsAnsiCount);

		if (m_FolderUserFieldsAnsiCount && !m_FieldDefinitionsA.empty())
		{
			auto i = 0;
			for (auto& fieldDefinition : m_FieldDefinitionsA)
			{
				auto szGUID = guid::GUIDToString(&fieldDefinition.Common.PropSetGuid);
				auto szFieldType = interpretprop::InterpretFlags(flagFolderType, fieldDefinition.FieldType);
				auto szFieldcap = interpretprop::InterpretFlags(flagFieldCap, fieldDefinition.Common.fcapm);

				szTmp = strings::formatmessage(
					IDS_FIELDANSIFIELD,
					i++,
					fieldDefinition.FieldType, szFieldType.c_str(),
					fieldDefinition.FieldNameLength,
					fieldDefinition.FieldName.c_str(),
					szGUID.c_str(),
					fieldDefinition.Common.fcapm, szFieldcap.c_str(),
					fieldDefinition.Common.dwString,
					fieldDefinition.Common.dwBitmap,
					fieldDefinition.Common.dwDisplay,
					fieldDefinition.Common.iFmt,
					fieldDefinition.Common.wszFormulaLength,
					fieldDefinition.Common.wszFormula.c_str());
				szFolderUserFieldStream += szTmp;
			}
		}

		szTmp = strings::formatmessage(
			IDS_FIELDUNICODEHEADER,
			m_FolderUserFieldsUnicodeCount);
		szFolderUserFieldStream += szTmp;

		if (m_FolderUserFieldsUnicodeCount && !m_FieldDefinitionsW.empty())
		{
			auto i = 0;
			for (auto& fieldDefinition : m_FieldDefinitionsW)
			{
				auto szGUID = guid::GUIDToString(&fieldDefinition.Common.PropSetGuid);
				auto szFieldType = interpretprop::InterpretFlags(flagFolderType, fieldDefinition.FieldType);
				auto szFieldcap = interpretprop::InterpretFlags(flagFieldCap, fieldDefinition.Common.fcapm);

				szTmp = strings::formatmessage(
					IDS_FIELDUNICODEFIELD,
					i++,
					fieldDefinition.FieldType, szFieldType.c_str(),
					fieldDefinition.FieldNameLength,
					fieldDefinition.FieldName.c_str(),
					szGUID.c_str(),
					fieldDefinition.Common.fcapm, szFieldcap.c_str(),
					fieldDefinition.Common.dwString,
					fieldDefinition.Common.dwBitmap,
					fieldDefinition.Common.dwDisplay,
					fieldDefinition.Common.iFmt,
					fieldDefinition.Common.wszFormulaLength,
					fieldDefinition.Common.wszFormula.c_str());
				szFolderUserFieldStream += szTmp;
			}
		}

		return szFolderUserFieldStream;
	}
}